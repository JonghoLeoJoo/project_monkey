"""단계별 실행 결과 캐싱 유틸리티.

캐시 파일 위치: .cache/{TICKER}_step{N}.json
  N=1: company_info
  N=2: company_info, facts, years
  N=3: company_info, facts, years, financial_data  (extract 직후)
  N=4: company_info, facts, years, financial_data  (market_cap/stock_prices 추가)
  N=5: company_info, facts, years, financial_data  (wacc_inputs 추가)

사용법:
  # 전체 실행 (매 단계 캐시 저장)
  python main.py AAPL

  # Step 4부터 재실행 (Step 3 캐시 로드)
  python main.py AAPL --from-step 4
"""

import json
from pathlib import Path
from typing import Any, Dict

CACHE_DIR = Path(__file__).parent.parent / '.cache'


# ---------------------------------------------------------------------------
# JSON 직렬화 헬퍼
# ---------------------------------------------------------------------------

def _stringify_keys(obj: Any) -> Any:
    """int 키를 str로 변환 (JSON 직렬화 전처리)."""
    if isinstance(obj, dict):
        return {str(k): _stringify_keys(v) for k, v in obj.items()}
    return obj


def _restore_int_keys(obj: Any) -> Any:
    """JSON 로드 후 순수 정수 문자열 키를 int로 복원 (회계연도 키 복원용)."""
    if isinstance(obj, dict):
        result = {}
        for k, v in obj.items():
            try:
                k = int(k)
            except (ValueError, TypeError):
                pass
            result[k] = _restore_int_keys(v)
        return result
    return obj


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def cache_path(ticker: str, step: int) -> Path:
    return CACHE_DIR / f"{ticker.upper()}_step{step}.json"


def save(ticker: str, step: int, state: Dict) -> None:
    """step 완료 직후 state를 캐시 파일에 저장한다."""
    CACHE_DIR.mkdir(exist_ok=True)
    dest = cache_path(ticker, step)
    with open(dest, 'w', encoding='utf-8') as f:
        json.dump(_stringify_keys(state), f, indent=2)
    print(f"  [cache] step {step} saved → .cache/{ticker.upper()}_step{step}.json")


def load(ticker: str, step: int) -> Dict:
    """step N 완료 상태를 캐시에서 로드한다.

    Args:
        ticker: 대문자 티커 심볼 (파일명 결정에 사용)
        step:   로드할 스텝 번호 (= from_step - 1)

    Raises:
        FileNotFoundError: 캐시 파일이 없을 때
    """
    p = cache_path(ticker, step)
    if not p.exists():
        raise FileNotFoundError(
            f"step {step} 캐시 파일을 찾을 수 없습니다: {p}\n"
            f"  먼저 전체 파이프라인을 실행해 캐시를 생성하세요 (--from-step 없이)."
        )
    with open(p, 'r', encoding='utf-8') as f:
        return _restore_int_keys(json.load(f))
