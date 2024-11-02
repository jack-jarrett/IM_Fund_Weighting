import pandas as pd
import openpyxl

def analyst_score(ar: str):
    # Transform analyst ratings into a score of 10
    if ar == "Strong buy":
        return 10
    elif ar == "Buy":
        return 7.5
    elif ar == "Neutral":
        return 5
    elif ar == "Sell":
        return 2.5
    elif ar == "Strong sell":
        return 0
    return 0

def rsi_score(rsi: float):
    # Transform RSI scores into a score of 10
    if rsi <= 30:
        return 10
    elif 30 < rsi <= 40:
        return 8
    elif 40 < rsi <= 50:
        return 6
    elif 50 < rsi <= 60:
        return 4
    elif 60 < rsi <= 70:
        return 2
    else:
        return 0

# filepath of the CSV downloaded from trading view
filepath = "/Users/jackjarrett/Downloads/Gambling Fund_2024-10-31.csv"

# Constants
PE_WEIGHT = 0.5
ANALYST_WEIGHT = 0.25
RSI_WEIGHT = 0.25
HOLDINGS_NUM = 5

# reading CSV into a dataframe and filtering for important columns
df = pd.read_csv(filepath)
df = df[['Symbol', 'Description', 'Price', 'Price - Currency', 'Market capitalization', 'Volume 1 day', 'Relative Strength Index (14) 1 day', 'Price to earnings ratio', 'Analyst Rating', 'Beta 5 years']]

# calculating the minimum and maximum PE ratios
min_pe = df["Price to earnings ratio"].min()
max_pe = df["Price to earnings ratio"].max()

# Collecting the total scores from PE Ratio, RSI Score, and Analyst Rating
df['PE_Score'] = 10 * (1 - (df['Price to earnings ratio'] - min_pe) / (max_pe - min_pe))
df['RSI_Score'] = df['Relative Strength Index (14) 1 day'].apply(rsi_score)
df['Analyst_Score'] = df['Analyst Rating'].apply(analyst_score)
df['Total Score'] = df['PE_Score']*PE_WEIGHT+df['Analyst_Score']*ANALYST_WEIGHT+df["RSI_Score"]*RSI_WEIGHT

# Sorting by total score and collecting the number of stocks specified by the HOLDINGS_NUM constant variable
df = df.sort_values(by=['Total Score'], ascending=False)
portfolio = df.head(HOLDINGS_NUM)

# Weighting the protfolio by total score / sum of total scores
total_of_scores = portfolio['Total Score'].sum()
portfolio['Weight'] = portfolio['Total Score'] / total_of_scores

# Printing the portfolio and its beta
print(portfolio)
print('Beta of the fund: ', portfolio['Beta 5 years'].sum() / portfolio.shape[0])

# saving the portfolio to an excel fund_holdings.xlsx
portfolio.to_excel("fund_holdings.xlsx", sheet_name="Portfolio")