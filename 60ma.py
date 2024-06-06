import pandas as pd
import requests
from bs4 import BeautifulSoup
from io import StringIO
import yfinance as yf
from datetime import datetime, timedelta
from telegram import Update, Bot
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes, JobQueue
import nest_asyncio
import asyncio

# nest_asyncio 적용
nest_asyncio.apply()

# 텔레그램 봇 설정
BOT_TOKEN = '7395106211:AAFiyE_uKx-X1SiH9WGnM_ndOplV07Le7nc'
CHAT_ID = '67612254'

# KOSPI 종목 코드와 종목명을 가져오는 함수
def get_kospi_stocks():
    url = 'https://kind.krx.co.kr/corpgeneral/corpList.do?method=download&marketType=stockMkt'
    response = requests.get(url)
    response.encoding = 'euc-kr'
    soup = BeautifulSoup(response.text, 'html.parser')
    table = soup.find('table')
    table_str = str(table)
    df = pd.read_html(StringIO(table_str), header=0)[0]
    df['종목코드'] = df['종목코드'].apply(lambda x: f"{x:06d}.KS")
    df = df[~df['회사명'].str.contains('스팩')]  # '스팩'이 포함된 종목 제외
    kospi_stocks = df[['종목코드', '회사명']].values.tolist()
    return kospi_stocks

# KOSDAQ 종목 코드와 종목명을 가져오는 함수
def get_kosdaq_stocks():
    url = 'https://kind.krx.co.kr/corpgeneral/corpList.do?method=download&marketType=kosdaqMkt'
    response = requests.get(url)
    response.encoding = 'euc-kr'
    soup = BeautifulSoup(response.text, 'html.parser')
    table = soup.find('table')
    table_str = str(table)
    df = pd.read_html(StringIO(table_str), header=0)[0]
    df['종목코드'] = df['종목코드'].apply(lambda x: f"{x:06d}.KQ")
    df = df[~df['회사명'].str.contains('스팩')]  # '스팩'이 포함된 종목 제외
    kosdaq_stocks = df[['종목코드', '회사명']].values.tolist()
    return kosdaq_stocks

# 주식 종목 체크 및 텔레그램 메시지 보내기
async def check_stocks(context: ContextTypes.DEFAULT_TYPE):
    today = datetime.now().date()
    end_date = today + timedelta(days=1)  # yfinance에서 end 날짜는 다음날로 설정해야 오늘까지 포함됨

    kospi_stocks = get_kospi_stocks()
    kosdaq_stocks = get_kosdaq_stocks()
    all_stocks = kospi_stocks + kosdaq_stocks

    crossed_stocks = []

    for ticker, name in all_stocks:
        data = yf.download(ticker, start='2020-01-01', end=end_date.strftime('%Y-%m-%d'))

        if data.empty:
            continue

        # 주봉 데이터 생성하기
        data['Week'] = data.index.to_period('W')
        weekly_data = data.resample('W').agg({'Open': 'first', 
                                              'High': 'max',
                                              'Low': 'min',
                                              'Close': 'last',
                                              'Volume': 'sum'})

        # 60주 이동평균선 계산하기
        weekly_data['60_MA'] = weekly_data['Close'].rolling(window=60).mean()

        if len(weekly_data) < 60:
            continue

        # 최근 주봉 데이터 가져오기
        latest_week = weekly_data.iloc[-1]
        previous_week = weekly_data.iloc[-2]

# 오늘 처음으로 종가가 60주선을 돌파했는지 확인하기
        if latest_week['Close'] > latest_week['60_MA'] and previous_week['Close'] <= previous_week['60_MA']:
            # 시가 총액 가져오기
            stock_info = yf.Ticker(ticker).info
            market_cap = stock_info.get('marketCap', 0)
            if market_cap >= 300000000000:
                close_price = int(latest_week['Close'])
                ma_60_price = int(latest_week['60_MA'])
                percentage = ((latest_week['Close'] - latest_week['60_MA']) / latest_week['60_MA']) * 100
                message = f"{name}: 종가 {close_price:,}원 / 60주선을 {percentage:.1f}% 돌파"
                crossed_stocks.append(message)
                if len(crossed_stocks) == 3:
                    break

    if crossed_stocks:
        full_message = "\n".join(crossed_stocks)
        await context.bot.send_message(chat_id=CHAT_ID, text=full_message)

async def main():
    # ApplicationBuilder를 사용하여 Application 객체 생성
    application = ApplicationBuilder().token(BOT_TOKEN).build()

    # JobQueue에 작업 추가
    job_queue = application.job_queue
    job_queue.run_once(check_stocks, 0)  # 즉시 주식 종목 체크 및 메시지 보내기

    # 봇 시작 및 폴링
    await application.initialize()
    await application.start()
    await application.stop()
    await application.shutdown()

if __name__ == '__main__':
    loop = asyncio.get_event_loop()
    loop.run_until_complete(main())


# 차트 그리기
'''
plt.figure(figsize=(14, 7))
plt.plot(weekly_data.index, weekly_data['Close'], label='Close Price')
plt.plot(weekly_data.index, weekly_data['60_MA'], label='60 Week MA', linestyle='--')
plt.xlabel('Date')
plt.ylabel('Price')
plt.title(f'{ticker} Stock Price and 60 Week Moving Average')
plt.legend()
plt.show()
'''