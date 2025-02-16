import pandas as pd
import matplotlib.pyplot as plt
import requests
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import random

# Constants
NEPSE_URL = "https://www.nepalstock.com/today-price"
WEATHER_URL = "https://api.openweathermap.org/data/2.5/weather"
QUOTES_URL = "https://zenquotes.io/api/quotes"
WEATHER_API_KEY = "50680f35bd3743853676d4fb3ddadc60"
CITY = "Kathmandu"
SENDER_EMAIL = "anjalisimkhada5@gmail.com'"
RECEIVER_EMAIL = "anjali.simkhada@westcliff.edu"
EMAIL_PASSWORD = "nvof bers oduz kdqd"  

# Part 1: Scrape Nepse Stock Data
def get_nepse_data():
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    try:
        driver.get(NEPSE_URL)
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".table.table__lg.table-striped.table__border.table__border--bottom"))
        )
        page_source = driver.page_source
        soup = BeautifulSoup(page_source, "html.parser")
        table = soup.find("table", class_="table table__lg table-striped table__border table__border--bottom")
        if table is None:
            print("No table was found.")
        else:
            data = []
            rows = table.find_all("tr")
            for row in rows[1:]:  # Skip header
                cols = row.find_all("td")
                if len(cols) > 1:
                    company_name = cols[1].get_text(strip=True)
                    stock_price = cols[2].get_text(strip=True).replace(',', '')
                    stock_price = float(stock_price)
                    data.append([company_name, stock_price])
            df = pd.DataFrame(data, columns=['Company', 'Stock Price'])
            df.to_csv('nepse_data.csv', index=False)
            print("Nepse data saved successfully.")
    except Exception as e:
        print(f"Error: {e}")
    finally:
        driver.quit()

# Part 2: Fetch Weather Data
def get_weather_data():
    params = {
        "q": CITY,
        "appid": WEATHER_API_KEY,
        "units": "metric"
    }
    weather_info = {}
    try:
        response = requests.get(WEATHER_URL, params=params)
        response.raise_for_status()
        weather_data = response.json()
        weather_info["city"] = weather_data["name"]
        weather_info["temperature"] = weather_data["main"]["temp"]
        weather_info["min_temp"] = weather_data["main"]["temp_min"]
        weather_info["feels_like"] = weather_data["main"]["feels_like"]
        weather_info["humidity"] = weather_data["main"]["humidity"]
        weather_info["weather_condition"] = weather_data["weather"][0]["description"]
        weather_info["wind_speed"] = weather_data["wind"]["speed"]
        weather_info["visibility"] = weather_data["visibility"]
    except Exception as e:
        print(f"Error occurred: {e}")
    return weather_info

# Part 3: Data Visualization (Matplotlib Chart)
def get_nepse_chart():
    df = pd.read_csv('nepse_data.csv')
    top_10_companies = df.sort_values(by='Stock Price', ascending=False).head(10)
    plt.figure(figsize=(10, 6))
    plt.bar(top_10_companies['Company'], top_10_companies['Stock Price'], color='skyblue')
    plt.title('Top 10 Companies by Stock Price (Nepse)', fontsize=14)
    plt.xlabel('Company', fontsize=12)
    plt.ylabel('Stock Price (NPR)', fontsize=12)
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    plt.savefig('nepse_chart.png')
    plt.show()

# Part 4: Fetch a Random Quote
def get_quote():
    try:
        response = requests.get(QUOTES_URL)
        response.raise_for_status()
        quotes = response.json()
        random_index = random.randint(0, len(quotes) - 1)  # Fixed: Added missing parenthesis
        quote = quotes[random_index]["q"]
        author = quotes[random_index]["a"]
    except Exception as e:
        print(f"Error: {e}")
        quote, author = "Stay positive!", "Unknown"
    return quote, author

# Part 5: Send Daily Email
def send_email(weather_info, quote, author):
    message = MIMEMultipart()
    message["From"] = SENDER_EMAIL
    message["To"] = RECEIVER_EMAIL
    message["Subject"] = "Your Daily Inspiration & Nepse Report"

    body = f"""
    Hello, this is your daily inspiration and Nepse report.

    Daily Inspiration:
    ----------------------
    "{quote}"
    - {author}

    Weather in Kathmandu:
    ----------------------
    City: {weather_info['city']}
    Temperature: {weather_info['temperature']}°C (Feels like {weather_info['feels_like']}°C)
    Minimum Temperature: {weather_info['min_temp']}°C
    Humidity: {weather_info['humidity']}%
    Weather Condition: {weather_info['weather_condition']}
    Wind Speed: {weather_info['wind_speed']} m/s
    Visibility: {weather_info['visibility']} meters

    Stay inspired and have a great day!
    """
    message.attach(MIMEText(body, "plain"))

    # Attach the Nepse chart
    file_path = 'nepse_chart.png'
    try:
        with open(file_path, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={file_path}")
            message.attach(part)
    except Exception as e:
        print(f"Error attaching file: {e}")

    # Send the email
    try:
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(SENDER_EMAIL, EMAIL_PASSWORD)
            server.sendmail(SENDER_EMAIL, RECEIVER_EMAIL, message.as_string())
        print("Email sent successfully!")
    except Exception as e:
        print(f"Error: {e}")

# Main Function
def main():
    # Scrape Nepse data
    get_nepse_data()

    # Fetch weather data
    weather_info = get_weather_data()

    # Generate Nepse chart
    get_nepse_chart()

    # Fetch a random quote
    quote, author = get_quote()

    # Send the email
    send_email(weather_info, quote, author)

if __name__ == "__main__":
    main()