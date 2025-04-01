import pandas as pd
from openpyxl import load_workbook

# Создает словарь тикеров из файла Excel
def create_ticker_dict(filepath):
    df_tickers = pd.read_excel(filepath, sheet_name='DDE')
    instrument_ticker = dict(zip(df_tickers['SHORTNAME'], df_tickers['CODE']))
    return instrument_ticker

# Загрузка данных для словаря тикеров
filepath = 'DDE_.xlsx'
instrument_ticker = create_ticker_dict(filepath)
print(instrument_ticker)

# Загрузка данных для создания 1-минутных OHLC свечей
df = pd.read_excel('DDE_1.xlsx', sheet_name='DDE')

# Преобразование времени в начало минуты
df['TIME'] = pd.to_datetime(df['Время'], format='%H:%M:%S').dt.floor('min')
df['TIME'] = df['TIME'].dt.strftime('%H:%M:%S')

# Добавление тикера и фильтрация
df['Ticker'] = df['Инструмент сокр.'].map(instrument_ticker)
df = df.dropna(subset=['Ticker'])

# Расчет OHLC данных
ohlc = df.groupby(['Ticker', 'TIME']).agg(
    OPEN=('Цена', 'first'),
    HIGH=('Цена', 'max'),
    LOW=('Цена', 'min'),
    CLOSE=('Цена', 'last')
).reset_index()

# Добавляем новые столбцы
ohlc['DIFF1'] = ohlc['HIGH'] - ohlc['OPEN']
ohlc['DIFF2'] = ohlc['OPEN'] - ohlc['LOW']
ohlc['DELTA'] = ohlc['DIFF1'] - ohlc['DIFF2']

# Сортировка и расчет кумулятивной суммы
ohlc = ohlc.sort_values(by=['Ticker', 'TIME'])
ohlc['CUM'] = ohlc.groupby('Ticker')['DELTA'].cumsum()

# Получаем последние значения CUM для каждого тикера
last_cum = ohlc.groupby('Ticker')['CUM'].last().reset_index()
last_cum.columns = ['Тикер', 'LAST CUM']

# Получаем последние значения CLOSE для каждого тикера
last_close = ohlc.groupby('Ticker')['CLOSE'].last().reset_index()
last_close.columns = ['Тикер', 'LAST CLOSE']

# Объединяем last_cum и last_close по тикеру
last_cum = pd.merge(last_cum, last_close, on='Тикер')

# Рассчитываем столбец PER
last_cum['PER'] = (last_cum['LAST CUM'] / last_cum['LAST CLOSE']) * 100

# Сортировка
last_cum = last_cum.sort_values(by='PER', ascending=False)

print(last_cum)
