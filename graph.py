import matplotlib.pyplot as plt
import pandas as pd

worksheet = pd.read_excel('./weight.xlsx')

col_values = worksheet.iloc[:, 0].dropna().values

pngDate = str(col_values[-1])
weight_png_date = pd.to_datetime(pngDate).date()

with plt.style.context('Solarize_Light2'):
    plt.rcParams["figure.figsize"] = (12, 6)
    plt.ylim(66, 84)

    plt.plot(worksheet['date'], worksheet['weight'], marker='o', color='#4e3b2f')
    plt.title('Weight Graph')
    plt.xlabel('Date')
    plt.ylabel('Weight')

    plt.xticks(worksheet['date'], rotation=45, ha='right')

    plt.grid(True)
    plt.tight_layout()
    plt.savefig('weight.png')
