When you run main.py two input requests will follow:
1) Enter company name or ticker: SEC listed company ticker or company name (e.g. AAPL or Apple)
2) Share price date for WACC (YYYY-MM-DD, or Enter for latest): XXXX-XX-XX

With the given information above, the main code wil generate an excel file containing the following worksheets:
- Financial Statements (past 4 years' actuals, followed by 5 years' projected financials following company's previous 5 year CAGR and operating ratios)
- WACC: calculations done based on share price of input date and 10 comparable companies' market cap and de-levered beta with re-levered beta, Cost of Equity and WACC calculation
- DCF model with FCF projections, using WACC, revenue growth from FS model inputs, terminal growth rate, with implied share price
