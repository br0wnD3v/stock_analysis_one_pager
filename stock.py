from openai import OpenAI
import requests
import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
import sys
import io
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from pdf2docx import parse

client = OpenAI(api_key='')

headers = {
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 14_0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.6668.101 Safari/537.36'
}


def get_stock_price(ticker):
    url = f'https://finance.yahoo.com/quote/{ticker}/key-statistics/'
    response = requests.get(url, headers=headers)
    soup = BeautifulSoup(response.content, 'lxml')
    price_element = soup.find('fin-streamer', {'data-field': 'regularMarketPrice'})
    stock_price = None
    
    if price_element:
        stock_price = price_element.get('data-value')
        return float(stock_price)
    
    return 0

"""P/E TTM"""
def get_trailing_pe(ticker):
    url = f'https://finance.yahoo.com/quote/{ticker}/key-statistics/'

    response = requests.get(url, headers=headers)
    soup = BeautifulSoup(response.content, 'lxml')
    rows = soup.find_all('tr')
    pe_value = None

    for row in rows:
        if 'Trailing P/E' in row.text:
            pe_value = row.find_all('td')[1].text
            break

    pe_value_float = 0
    if pe_value:
        try:
            pe_value_float = float(pe_value.replace(',', ''))  # Convert to float and handle commas
        except ValueError:
            pe_value_float = 0
    else:
        print(f"Could not find Trailing P/E for {ticker}")
        return 0
    
    return pe_value_float

"""EV/EBITDA"""
def get_ev_ebitda(ticker):
    url = f'https://finance.yahoo.com/quote/{ticker}/key-statistics/'

    response = requests.get(url, headers=headers)
    soup = BeautifulSoup(response.content, 'lxml')
    rows = soup.find_all('tr')

    ev_ebitda_value = None

    for row in rows:
        if 'Enterprise Value/EBITDA' in row.text:
            ev_ebitda_value = row.find_all('td')[1].text
            break

    # Check if the value is found
    ev_ebitda_value_float = 0
    if ev_ebitda_value:
        try:
            ev_ebitda_value_float = float(ev_ebitda_value.replace(',', ''))  # Convert to float and handle commas
        except ValueError:
            ev_ebitda_value_float = 0
    else :
        print(f"Could not find EV/EBITA for {ticker}")
        return 0

    return ev_ebitda_value_float
    

"""P/S TTM"""
def get_price_sales(ticker):
    url = f'https://finance.yahoo.com/quote/{ticker}/key-statistics/'
    response = requests.get(url, headers=headers)
    soup = BeautifulSoup(response.content, 'lxml')
    rows = soup.find_all('tr')

    ps_value = None

    for row in rows:
        if 'Price/Sales' in row.text:
            ps_value = row.find_all('td')[1].text
            break

    ps_value_float=0
    if ps_value:
        try:
            ps_value_float = float(ps_value.replace(',', ''))  # Convert to float and handle commas
        except ValueError:
            ps_value_float = 0
    else:
        print(f"Could not find P/S TTM for {ticker}")
        return 0
    
    return ps_value_float

"""Dividend Yield"""
# Loop through each ticker symbol in the list
def get_divident_yield(ticker):
        stock = yf.Ticker(ticker)
        info = stock.info

        # Get the Forward Dividend Yield information
        dividend_yield = info.get("dividendYield")  # As a decimal (e.g., 0.05 for 5%)

        # Display the Forward Dividend Yield in the specified format
        if dividend_yield:
            yield_percentage = dividend_yield * 100
        else:
            print(f"Could not find DY for {ticker}")
            return 0
        
        return yield_percentage

# Function to extract Price/Book for a given ticker
"""P/NAV OR P/B TTM"""
def get_price_book(ticker):
    url = f'https://finance.yahoo.com/quote/{ticker}/key-statistics/'

    response = requests.get(url, headers=headers)
    soup = BeautifulSoup(response.content, 'lxml')
    rows = soup.find_all('tr')

    pb_value = None

    for row in rows:
        if 'Price/Book' in row.text:
            pb_value = row.find_all('td')[1].text
            break

    if pb_value:
        try:
            pb_value_float = float(pb_value.replace(',', ''))  # Convert to float and handle commas
        except ValueError:
            pb_value_float = 0
    else:
        print(f"Could not find P/NAV for {ticker}")
        return 0
    
    return pb_value_float   


def convert_pdf2docx(input_file: str, output_file: str):
    parse(input_file, output_file)

def format_market_value(value):
    if value is None:
        return "Not available"

    value_bil = value / 1e9
    return f"${value_bil:.2f}B"

def format_date(date):
    if date is None:
        return "Not available"
    try:
        if not isinstance(date, datetime):
            date = datetime.fromtimestamp(date)
        return date.strftime('%d-%m-%Y')
    except:
        return "Not available"

def get_company_info(stock_info):
    try:
        print("\nCOMPANY OVERVIEW")
        print(f"Company Name: {stock_info.get('longName', 'Not available')}")
        print(f"Industry: {stock_info.get('industry', 'Not available')}")
        print(f"Sector: {stock_info.get('sector', 'Not available')}")

        if stock_info.get('longBusinessSummary'):
            sentences = stock_info['longBusinessSummary'].split('. ')
            summary = '. '.join(sentences[:3]) + '.'
            print("\nBusiness Description:")
            print(summary)

    except Exception as e:
        print(f"Error getting company info: {e}")


def get_business_details(stock_info):
    try:
        print("\nBUSINESS AND MARKET POSITION")
        if stock_info.get('marketCap'):
            print("\nMarket Position:")
            print(f"• Market Cap: ${stock_info['marketCap']/1e9:.1f}B")
        if stock_info.get('sharesOutstanding'):
            print(f"• Shares Outstanding: {stock_info['sharesOutstanding']/1e6:.1f}M")
        if stock_info.get('floatShares'):
            print(f"• Float: {stock_info['floatShares']/1e6:.1f}M")

        print("\nKey Statistics:")

    except Exception as e:
        print(f"Error getting business details: {e}")

def get_investment_thesis(stock_info):
    try:
        print("\nINVESTMENT THESIS")

        print("\nFinancial Health:")
        if stock_info.get('totalCash'):
            print(f"• Cash Position: ${stock_info['totalCash']/1e9:.1f}B")
        if stock_info.get('totalDebt'):
            print(f"• Total Debt: ${stock_info['totalDebt']/1e9:.1f}B")
        if stock_info.get('debtToEquity'):
            print(f"• Debt to Equity: {stock_info['debtToEquity']:.2f}")
        if stock_info.get('currentRatio'):
            if(stock_info.get('currentRatio') < 1):
                print(f"• Current Ratio: {stock_info['currentRatio']:.2f}")
            else :
                print(f"• Current Ratio: {stock_info['currentRatio']:.1f}")


        print("\nAnalyst Insights:")
        if stock_info.get('recommendationMean'):
            print(f"• Analyst Rating (1-5): {stock_info['recommendationMean']:.1f}")
        if stock_info.get('recommendationKey'):
            print(f"• Recommendation: {stock_info['recommendationKey'].upper()}")
        if stock_info.get('numberOfAnalystOpinions'):
            print(f"• Number of Analysts: {stock_info['numberOfAnalystOpinions']}")
        if stock_info.get('targetMeanPrice'):
            print(f"• Mean Target Price: ${stock_info['targetMeanPrice']:.1f}")
            current_price = stock_info.get('currentPrice', stock_info.get('regularMarketPrice'))
            if current_price:
                upside = ((stock_info['targetMeanPrice'] / current_price) - 1) * 100
                print(f"• Implied +/-: {upside:.1f}%")

    except Exception as e:
        print(f"Error getting investment thesis: {e}")

def create_financial_table(income_stmt, balance_sheet):
    try:
        if income_stmt is not None and not income_stmt.empty:
            financial_data = pd.DataFrame()

            def convert_to_millions(x):
                try:
                    return float(x) / 1_000_000
                except:
                    return float('nan')

            try:
                financial_data['Revenue'] = income_stmt.loc['Total Revenue'].apply(convert_to_millions)
            except:
                try:
                    financial_data['Revenue'] = income_stmt.loc['Revenue'].apply(convert_to_millions)
                except:
                    financial_data['Revenue'] = float('nan')

            try:
                financial_data['EBIT'] = income_stmt.loc['Operating Income'].apply(convert_to_millions)
            except:
                try:
                    financial_data['EBIT'] = income_stmt.loc['EBIT'].apply(convert_to_millions)
                except:
                    financial_data['EBIT'] = float('nan')

            try:
                financial_data['Net Profit'] = income_stmt.loc['Net Income'].apply(convert_to_millions)
            except:
                try:
                    financial_data['Net Profit'] = income_stmt.loc['Net Income Common Stockholders'].apply(convert_to_millions)
                except:
                    financial_data['Net Profit'] = float('nan')

            try:
                ebit = income_stmt.loc['Operating Income']
                try:
                    depreciation = income_stmt.loc['Depreciation & Amortization']
                except:
                    depreciation = income_stmt.loc['Depreciation And Amortization']

                financial_data['EBITDA'] = (ebit + depreciation).apply(convert_to_millions)
            except:
                try:
                    financial_data['EBITDA'] = income_stmt.loc['EBITDA'].apply(convert_to_millions)
                except:
                    financial_data['EBITDA'] = float('nan')

            try:
                net_profit = income_stmt.loc['Net Income']
                total_assets = balance_sheet.loc['Total Assets']
                roi = (net_profit / total_assets) * 100
                financial_data['ROI'] = roi.apply(lambda x: f"{x:.2f}%")
            except:
                financial_data['ROI'] = "N/A"

            financial_data.index = financial_data.index.year
            financial_data = financial_data.sort_index(ascending=False)

            pd.set_option('display.float_format', lambda x: '{:,.0f}'.format(x))

            print("\nFINANCIAL TABLE (in millions USD)")

            print(financial_data.fillna('N/A'))

            return financial_data

    except Exception as e:
        print(f"Error creating financial table: {e}")
        return None

def get_risk_analysis(ticker):
    try:
        print("\nRISK ANALYSIS AND MITIGATION")
        content = f'What are the top 2 risks and mitigations for ${ticker} stock? Provide two-liner explanation of the point as well. Use - at the start of each line. Dont bold anything anywhere.'
        completion = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {
                   "role": "user",
                    "content": [
                        {
                            "type":"text",
                            "text":content
                        }
                    ]
                }
            ]
        )
        print(completion.choices[0].message.content)

    except Exception as e:
        print(f"Error generating risk analysis: {e}")

def get_stock_strengths(ticker):
    try:
        print("\nKEY STRENGTHS")
        content = f'${ticker}stock and company strengths(top 3)? Provide one-line explanation of the point as well. Use - at the start of each line. dont use bold anywhere.'
        completion = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {
                   "role": "user",
                    "content": [
                        {
                            "type":"text",
                            "text":content
                        }
                    ]
                }
            ]
        )
        print(completion.choices[0].message.content)
       
    except Exception as e:
        print("Error analyzing strengths:", e)

def get_stock_catalysts(ticker):
    try:
        print("\nGROWTH CATALYSTS")
        # strengths = []
        content = f'${ticker}stock and company growth catalysts(top 3)? Provide one-line explanation of the point as well. Use - at the start of each line. dont use bold anywhere.'
        completion = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {
                   "role": "user",
                    "content": [
                        {
                            "type":"text",
                            "text":content
                        }
                    ]
                }
            ]
        )
        print(completion.choices[0].message.content)

    except Exception as e:
        print("Error analyzing catalysts:", e)


def get_assets_under_management_ratio(ticker):
    try:
        # strengths = []
        content = f'Only give one word answers, Whats ${ticker} Assets Under management numeric value. Without commas or $ sign.'
        completion = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {
                   "role": "user",
                    "content": [
                        {
                            "type":"text",
                            "text":content
                        }
                    ]
                }
            ]
        )
        stock_price = get_stock_price(ticker)
        aum = int(completion.choices[0].message.content)
        aum = aum/1000000000    
        
        ratio = stock_price/aum
        return ratio


    except Exception as e:
        return 0


def get_metric_color(metric_name, stock_value, comparable_value):
    """
    Determine the color for a metric based on comparison with comparable companies
    Returns hex color value
    """
    if stock_value == 0 or comparable_value == 0:
        return '#000000'  # black
    
    percentage_diff = ((stock_value - comparable_value) / comparable_value) * 100
    
    if metric_name == "Dividend Yield":
        # Higher dividend yield is better
        if percentage_diff > 10:
            return '#008000'  # green
        elif percentage_diff >= -10:
            return '#FFA500'  # orange
        else:
            return '#FF0000'  # red
    else:
        # Lower values are better for other metrics
        if percentage_diff < 0:
            return '#008000'  # green
        elif percentage_diff <= 10:
            return '#FFA500'  # orange
        else:
            return '#FF0000'  # red

def format_metric_value(metric_name, value):
    """
    Format metric values with appropriate decimal places
    """
    if metric_name == "Dividend Yield":
        return f"{value:.2f}%"
    else:
        return f"{value:.2f}"

def get_metric_value(ticker, metric_name):
    """
    Get metric value from Yahoo Finance
    """
    if ticker == 'PRIVATE' or ticker == 'Private' or ticker.lower() == 'private':
        return 0
    else :
        if metric_name == "EV/EBITDA":
            return get_ev_ebitda(ticker)
        elif metric_name == "P/E TTM":
            return get_trailing_pe(ticker)
        elif metric_name == "P/NAV":
            return get_price_book(ticker)
        elif metric_name == "Dividend Yield":
            return get_divident_yield(ticker)
        elif metric_name == "P/S TTM":
            return get_price_sales(ticker)
        elif metric_name == 'P/AUM':
            return get_assets_under_management_ratio(ticker)
        return 0

def create_pdf_report(ticker, all_content, final_metrics, comparable_metrics, comparable_debt_equity, comparable_current_ratio, comparable_upside):
    try:
        filename = f"{ticker}_Analysis_{datetime.now().strftime('%Y%m%d')}.pdf"
        margins = (15, 10, 10, 10)
        doc = SimpleDocTemplate(filename, pagesize=letter, leftMargin=margins[0], rightMargin=margins[1], topMargin=margins[2], bottomMargin=margins[3])
        styles = getSampleStyleSheet()
        story = []
        normal_style = ParagraphStyle('CustomNormal', parent=styles['Normal'], fontSize=9, spaceAfter=4, spaceBefore=4, fontName='Helvetica', leading=11)
        heading_style = ParagraphStyle('CustomHeading', parent=styles['Normal'], fontSize=10, fontName='Helvetica-Bold', textColor=colors.maroon, spaceAfter=4, spaceBefore=8, leading=12)
        bold_style = ParagraphStyle('CustomBold', parent=styles['Normal'], fontSize=9, spaceAfter=4, spaceBefore=4, fontName='Helvetica-Bold', leading=11)

        def get_ratio_color(metric_name, stock_value, comparable_value):
            """
            Determine color for financial ratios (Debt/Equity, Current Ratio, Implied Upside)
            Returns hex color value
            """
            if stock_value == 0 or comparable_value == 0:
                return '#000000'  # black
            
            percentage_diff = ((stock_value - comparable_value) / comparable_value) * 100
            
            if metric_name == "Implied +/-":
                # Higher upside is better
                if percentage_diff > 10:
                    return '#008000'  # green
                elif percentage_diff >= -10:
                    return '#FFA500'  # orange
                else:
                    return '#FF0000'  # red
            elif metric_name == "Current Ratio":
                    if 1.0 <= stock_value <= 1.5:
                        return '#008000'  # green
                    elif 0.8 <= stock_value < 1.0 or 1.5 < stock_value <= 2.0:
                        return '#FFA500'  # orange
                    else:
                        return '#FF0000'  # red
                    
            elif metric_name == 'P/AUM':
                    if(percentage_diff < 0):
                        return '#008000'  # orange    
                    elif 0 < percentage_diff < 10:
                        return '#FFA500'  # orange
                    else :
                        return '#FF0000'  # red

            else:  # Debt/Equity
                    if percentage_diff < -10:
                        return '#008000'  # green
                    elif percentage_diff <= 10:
                        return '#FFA500'  # orange
                    else:
                        return '#FF0000'  # red

        # Capture different sections
        overview_output = io.StringIO()
        sys.stdout = overview_output
        get_company_info(all_content['info'])
        overview_content = overview_output.getvalue()

        business_output = io.StringIO()
        sys.stdout = business_output
        get_business_details(all_content['info'])
        business_content = business_output.getvalue()

        financial_output = io.StringIO()
        sys.stdout = financial_output
        create_financial_table(all_content['income_stmt'], all_content['balance_sheet'])
        financial_content = financial_output.getvalue()

        strengths_output = io.StringIO()
        sys.stdout = strengths_output
        get_stock_strengths(ticker)
        strengths_content = strengths_output.getvalue()

        catalysts_output = io.StringIO()
        sys.stdout = catalysts_output
        get_stock_catalysts(ticker)
        catalysts_content = catalysts_output.getvalue()

        analysis_output = io.StringIO()
        sys.stdout = analysis_output
        get_investment_thesis(all_content['info'])
        get_risk_analysis(ticker)
        analysis_content = analysis_output.getvalue()

        events_output = io.StringIO()
        sys.stdout = events_output
        if all_content['info'].get('earningsDate'):
            print("\nUPCOMING EVENTS")
            if isinstance(all_content['info']['earningsDate'], list) and len(all_content['info']['earningsDate']) > 0:
                earnings_date = format_date(all_content['info']['earningsDate'][0])
                print(f"Next Earnings Date: {earnings_date}")
        if all_content['info'].get('exDividendDate'):
            div_date = format_date(all_content['info']['exDividendDate'])
            print(f"Ex-Dividend Date: {div_date}")
            if all_content['info'].get('dividendRate'):
                print(f"Dividend Rate: ${all_content['info']['dividendRate']:.2f}")
            if all_content['info'].get('dividendYield'):
                print(f"Dividend Yield: {all_content['info']['dividendYield']*100:.2f}%")
        events_content = events_output.getvalue()

        sys.stdout = sys.__stdout__

        def process_section(content, story, bullets=False):
            lines = content.split('\n')
            headers = ["COMPANY OVERVIEW", "BUSINESS AND MARKET POSITION", "FINANCIAL TABLE", "KEY STRENGTHS", 
                      "GROWTH CATALYSTS", "INVESTMENT THESIS", "RISK ANALYSIS AND MITIGATION", "UPCOMING EVENTS", "PRICE CHART"]
            bold_lines = ["Company Name:", "Industry:", "Sector:", "Business Description:", "Market Position:", 
                         "Key Statistics:", "Financial Health:", "Analyst Insights:", "Market & Competition Risks:", 
                         "1. Market Sensitivity Risk", "2. Industry Competition Risk", "Ex-Dividend Date:", 
                         "Dividend Rate:", "Dividend Yield:"]
            
            for line in lines:
                if line.strip():
                    if any(header in line.strip() for header in headers):
                        header_text = next(header for header in headers if header in line.strip())
                        story.append(Paragraph(header_text, heading_style))
                    elif any(bold_line in line.strip() for bold_line in bold_lines):
                        if ':' in line:
                            key, value = line.split(":", 1)
                            if bullets:
                                story.append(Paragraph(f'\u00A0\u00A0• <b>{key.strip()}</b>: {value.strip()}', normal_style))
                            else:
                                story.append(Paragraph(f'<b>{key.strip()}</b>: {value.strip()}', normal_style))
                    else:
                        story.append(Paragraph(line, normal_style))

        # Process overview section with chart
        overview_content_list = []
        process_section(overview_content, overview_content_list)

        # Create stock price chart
        plt.figure(figsize=(4, 3))
        history_data = all_content['history_data']
        plt.plot(history_data.index, history_data['Close'], label='Close Price', color='blue')
        history_data['MA50'] = history_data['Close'].rolling(window=50).mean()
        history_data['MA200'] = history_data['Close'].rolling(window=200).mean()
        plt.plot(history_data.index, history_data['MA50'], label='50-day MA', color='orange', linestyle='--')
        plt.plot(history_data.index, history_data['MA200'], label='200-day MA', color='red', linestyle='--')
        plt.title(f'{ticker} Stock Price - Last 12 Months', fontsize=8)
        plt.xlabel('Date', fontsize=8)
        plt.ylabel('Price (USD)', fontsize=8)
        plt.tick_params(axis='both', labelsize=6)
        plt.grid(True)
        plt.legend(fontsize=6)
        plt.tight_layout()

        img_data = io.BytesIO()
        plt.savefig(img_data, format='png', dpi=256, bbox_inches='tight')
        img_data.seek(0)
        plt.close()
        
        chart_img = Image(img_data)
        chart_img.drawHeight = 2.3*inch
        chart_img.drawWidth = 3.5*inch

        # Create side-by-side layout for overview and chart
        left_content = []
        left_content.extend(overview_content_list)
        left_wrapper = Table([[left_content]], colWidths=[doc.width/2.0 - 20])
        left_wrapper.setStyle(TableStyle([('LEFTPADDING', (0, 0), (-1, -1), -15), 
                                        ('RIGHTPADDING', (0, 0), (-1, -1), 10)]))

        right_content = []
        right_content.append(chart_img)
        right_wrapper = Table([[right_content]], colWidths=[doc.width/2.0 - 20])
        right_wrapper.setStyle(TableStyle([('LEFTPADDING', (0, 0), (-1, -1), 0), 
                                         ('RIGHTPADDING', (0, 0), (-1, -1), 10),
                                         ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                                         ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')]))

        overview_data = [[left_wrapper, right_wrapper]]
        overview_table = Table(overview_data, colWidths=[doc.width/2.0 - 12, doc.width/2.0 - 12])
        overview_table.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'TOP'),
                                          ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                                          ('LEFTPADDING', (0, 0), (-1, -1), 10),
                                          ('RIGHTPADDING', (0, 0), (-1, -1), 10)]))

        story.append(overview_table)
        story.append(Spacer(1, 6))

        # Process business section with financial table
        left_content = []
        left_content.append(Paragraph("BUSINESS AND MARKET POSITION", heading_style))

        business_lines = business_content.split('\n')
        for i, line in enumerate(business_lines):
            if line.strip():
                if not "BUSINESS AND MARKET POSITION" in line:
                    if "Key Statistics:" in line:
                        left_content.append(Paragraph(line, bold_style))
                        for idx, metric in enumerate(final_metrics):
                            stock_value = get_metric_value(ticker, metric)
                            comp_value = comparable_metrics[idx]
                            color_hex = get_metric_color(metric, stock_value, comp_value)
                            metric_text = f"• {metric}: "
                            stock_value_text = format_metric_value(metric, stock_value)
                            comp_value_text = format_metric_value(metric, comp_value)
                            paragraph = Paragraph(
                                f"{metric_text}"f'<font color="{color_hex}">{stock_value_text}</font>'
                                f" (Peer avg: {comp_value_text})", 
                                normal_style
                            )
                            left_content.append(paragraph)
                    elif any(bold_line in line.strip() for bold_line in ["Market Position:", "Business Performance:"]):
                        left_content.append(Paragraph(line, bold_style))
                    else:
                        left_content.append(Paragraph(line, normal_style))

        left_wrapper = Table([[left_content]], colWidths=[doc.width/2.0 - 20])
        left_wrapper.setStyle(TableStyle([('LEFTPADDING', (0, 0), (-1, -1), -15),
                                        ('RIGHTPADDING', (0, 0), (-1, -1), 10)]))

        financial_table_content = []
        financial_table_content.append(Paragraph("FINANCIAL TABLE (in millions USD)", heading_style))
        financial_lines = [line.strip() for line in financial_content.split('\n') if line.strip()]
        
        table_data = []
        headers = ['Year', 'Revenue', 'EBIT', 'Net Profit', 'EBITDA', 'ROI']
        table_data.append(headers)
        
        for line in financial_lines:
            parts = line.split()
            if parts and len(parts) >= 6 and parts[0].startswith('202'):
                table_data.append(parts)

        financial_table = Table(table_data)
        financial_table.setStyle(TableStyle([
            ('FONT', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('BACKGROUND', (0, 0), (-1, 0), colors.maroon),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('BOX', (0, 0), (-1, -1), 1, colors.black),
            ('LINEBELOW', (0, 0), (-1, 0), 1, colors.black),
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('ROWPADDING', (0, 0), (-1, -1), 4)
        ]))

        right_content = []
        right_content.extend(financial_table_content)
        right_content.append(financial_table)
        right_wrapper = Table([[right_content]], colWidths=[doc.width/2.0 - 20])
        right_wrapper.setStyle(TableStyle([('LEFTPADDING', (0, 0), (-1, -1), 0),
                                         ('RIGHTPADDING', (0, 0), (-1, -1), 10)]))

        side_by_side_data = [[left_wrapper, right_wrapper]]
        side_by_side = Table(side_by_side_data, colWidths=[doc.width/2.0 - 12, doc.width/2.0 - 12])
        side_by_side.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'TOP'),
                                        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                                        ('LEFTPADDING', (0, 0), (-1, -1), 10),
                                        ('RIGHTPADDING', (0, 0), (-1, -1), 10)]))

        story.append(Spacer(1, 6))
        story.append(side_by_side)
        story.append(Spacer(1, 6))

        # Process strengths and catalysts sections
        strengths_content_list = []
        process_section(strengths_content, strengths_content_list)
        left_wrapper = Table([[strengths_content_list]], colWidths=[doc.width/2.0 - 20])
        left_wrapper.setStyle(TableStyle([('LEFTPADDING', (0, 0), (-1, -1), -15),
                                        ('RIGHTPADDING', (0, 0), (-1, -1), 10)]))

        catalysts_content_list = []
        process_section(catalysts_content, catalysts_content_list)
        right_wrapper = Table([[catalysts_content_list]], colWidths=[doc.width/2.0 - 20])
        right_wrapper.setStyle(TableStyle([('LEFTPADDING', (0, 0), (-1, -1), 0),
                                         ('RIGHTPADDING', (0, 0), (-1, -1), 10)]))

        insights_data = [[left_wrapper, right_wrapper]]
        insights_table = Table(insights_data, colWidths=[doc.width/2.0 - 12, doc.width/2.0 - 12])
        insights_table.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'TOP'),
                                          ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                                          ('LEFTPADDING', (0, 0), (-1, -1), 10),
                                          ('RIGHTPADDING', (0, 0), (-1, -1), 10)]))

        story.append(insights_table)
        story.append(Spacer(1, 6))

        # Process investment thesis and risk analysis
        risk_content_list = []
        financial_insights = []
        analyst_insights = []
        
        analysis_lines = analysis_content.split('\n')
        current_section = None
        
        for line in analysis_lines:
            if "INVESTMENT THESIS" in line:
                current_section = "thesis"
            elif "Financial Health:" in line:
                current_section = "financial"
                financial_insights.append(Paragraph(line, bold_style))
            elif "Analyst Insights:" in line:
                current_section = "analyst"
                analyst_insights.append(Paragraph(line, bold_style))
            elif "RISK ANALYSIS" in line:
                current_section = "risk"
            elif line.strip():
                if current_section == "financial":
                    if "Financial Health:" in line:
                        continue
                    elif "Debt to Equity:" in line:
                        stock_de = float(line.split(": ")[1])
                        color_hex = get_ratio_color("Debt to Equity", stock_de, comparable_debt_equity)
                        financial_insights.append(Paragraph(
                            f"• Debt to Equity: <font color='{color_hex}'>{stock_de:.2f}</font> (Peer avg: {comparable_debt_equity:.2f})", 
                            normal_style
                        ))
                    elif "Current Ratio:" in line:
                        stock_cr = float(line.split(": ")[1])
                        color_hex = get_ratio_color("Current Ratio", stock_cr, comparable_current_ratio)
                        financial_insights.append(Paragraph(
                            f"• Current Ratio: <font color='{color_hex}'>{stock_cr:.2f}</font> (Peer avg: {comparable_current_ratio:.2f})", 
                            normal_style
                        ))
                    else:
                        financial_insights.append(Paragraph(line, normal_style))
                elif current_section == "analyst":
                    if "Analyst Insights:" in line:
                        continue
                    elif "Implied +/-:" in line:
                        try:
                            stock_upside = float(line.split(": ")[1].rstrip("%"))
                            color_hex = get_ratio_color("Implied +/-", stock_upside, comparable_upside)
                            analyst_insights.append(Paragraph(
                                f"• Implied +/-: <font color='{color_hex}'>{stock_upside:.1f}%</font> (Peer avg: {comparable_upside:.1f}%)", 
                                normal_style
                            ))
                        except ValueError:
                            analyst_insights.append(Paragraph(line, normal_style))
                    else:
                        analyst_insights.append(Paragraph(line, normal_style))
                elif current_section == "risk":
                    if "RISK ANALYSIS" not in line:
                        if any(bold_line in line.strip() for bold_line in ["Market & Competition Risks:", 
                                                                         "1. Market Sensitivity Risk",
                                                                         "2. Industry Competition Risk"]):
                            risk_content_list.append(Paragraph(line, bold_style))
                        else:
                            risk_content_list.append(Paragraph(line, normal_style))

        left_analysis = []
        left_analysis.append(Paragraph("INVESTMENT THESIS", heading_style))
        
        financial_wrapper = Table([[financial_insights]], colWidths=[doc.width/4.0 - 20])
        financial_wrapper.setStyle(TableStyle([('LEFTPADDING', (0, 0), (-1, -1), -5),
                                             ('RIGHTPADDING', (0, 0), (-1, -1), 10)]))
        
        analyst_wrapper = Table([[analyst_insights]], colWidths=[doc.width/4.0 - 20])
        analyst_wrapper.setStyle(TableStyle([('LEFTPADDING', (0, 0), (-1, -1), 0),
                                           ('RIGHTPADDING', (0, 0), (-1, -1), 10)]))
        
        insights_data = [[financial_wrapper, analyst_wrapper]]
        insights_table = Table(insights_data, colWidths=[doc.width/4.0, doc.width/4.0])
        insights_table.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'TOP'),
                                          ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                                          ('LEFTPADDING', (0, 0), (-1, -1), 10),
                                          ('RIGHTPADDING', (0, 0), (-1, -1), 10)]))
        
        left_analysis.append(insights_table)

        events_content_list = []
        process_section(events_content, events_content_list, True)
        left_wrapper = Table([[left_analysis], [events_content_list]], colWidths=[doc.width/2.0 - 20])
        left_wrapper.setStyle(TableStyle([('LEFTPADDING', (0, 0), (-1, -1), -5),
                                        ('RIGHTPADDING', (0, 0), (-1, -1), 10)]))

        right_analysis = []
        right_analysis.append(Paragraph("RISK ANALYSIS AND MITIGATION", heading_style))
        right_analysis.extend(risk_content_list)
        
        right_wrapper = Table([[right_analysis]], colWidths=[doc.width/2.0 - 20])
        right_wrapper.setStyle(TableStyle([('LEFTPADDING', (0, 0), (-1, -1), 0),
                                         ('RIGHTPADDING', (0, 0), (-1, -1), 10)]))
        
        analysis_data = [[left_wrapper, right_wrapper]]
        analysis_table = Table(analysis_data, colWidths=[doc.width/2.0, doc.width/2.0])
        analysis_table.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'TOP'),
                                          ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                                          ('LEFTPADDING', (0, 0), (-1, -1), 10),
                                          ('RIGHTPADDING', (0, 0), (-1, -1), 10)]))

        story.append(Spacer(1, 2))
        story.append(analysis_table)

        doc.build(story)
        print(f"\nPDF report saved as: {filename}")

        docx_name = f"{ticker}_Analysis_{datetime.now().strftime('%Y%m%d')}.docx"
        convert_pdf2docx(filename, docx_name)

    except Exception as e:
        print(f"Error creating PDF report: {e}")
        print(f"Error details: {str(e)}")

def extract_ticker(s):
    try:
        s = s.strip('"')
        parts = s.split("': '")
        if len(parts) == 2:
            # Return just the ticker part, removing the trailing quote
            return parts[1].rstrip("'")
        else:
            print(f"Warning: Unexpected format in string: {s}")
            return s
    except Exception as e:
        print(f"Error processing string: {s}")
        print(f"Error details: {e}")
        return s

def get_dividend_history(ticker):
    url = f"https://api.nasdaq.com/api/quote/{ticker}/dividends?assetclass=stocks"
    headers = {
        'User-Agent': 'Mozilla/5.0',
        'Accept': 'application/json',
        'Referer': 'https://www.nasdaq.com/'
    }
    
    response = requests.get(url, headers=headers)
    data = response.json()
    filename=""
    if data['data']:
        df = pd.DataFrame(data['data']['dividends']['rows'])
        filename = f'{ticker}_dividend_history.txt'
        df.to_csv(filename, sep='|', index=False)
        return filename
    return ""
    
def get_fed_rate():
    url = "https://fred.stlouisfed.org/series/FEDFUNDS"
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)',
        'Accept': 'text/html'
    }
    
    response = requests.get(url, headers=headers)
    soup = BeautifulSoup(response.text, 'html.parser')
    rate = soup.find('span', {'class': 'series-meta-observation-value'}).text
    return float(rate)

def calculate_dividend_growth(filename):
   df = pd.read_csv(filename, sep='|')
   df['amount'] = df['amount'].str.replace('$', '').astype(float)
   
   recent_sum = df['amount'][:4].sum()
   previous_sum = df['amount'][4:8].sum()
   dividend_growth = recent_sum - previous_sum
   growth_rate = int((dividend_growth / previous_sum) * 100 * 100) / 100   
   return growth_rate
   

def get_GGM(forward_dividend,fed_rate,beta,spy_total_return, dividend_growth) :
    expected_price = forward_dividend/(fed_rate + (beta*(spy_total_return - fed_rate))-dividend_growth)
    return expected_price

def main():
        try:
            ticker = input("\nEnter Stock Ticker (or 'quit' to exit): ").upper()
            print("\nFetching data...")
            stock = yf.Ticker(ticker)
            file_path = '35stocks.xlsx'
            df = pd.read_excel(file_path, sheet_name='Sheet1')
            df_transposed = df.T
            df_transposed.columns = ["Stock", "Metric_1", "Metric_2"]
            df_transposed.dropna(how="all", inplace=True)
            df_transposed.reset_index(drop=True, inplace=True)
            metrics=[]
            final_metrics= []
            for index, row in df_transposed.iterrows():
                stock_name = row['Stock']
                metrics = [row['Metric_1'], row['Metric_2']]
                if stock_name == ticker:
                    final_metrics = metrics
                    break

            print(final_metrics)

            file_path = '35stockscomparable.xlsx'
            df = pd.read_excel(file_path, sheet_name='Sheet1')
            original_columns = df.columns.tolist()
            comparable_companies = []
            if ticker in original_columns:
                column_data = df[ticker].iloc[0:].dropna().tolist()  # Changed from iloc[1:] to iloc[0:]
                comparable_companies = []
                for company in column_data:
                    try:
                        ticker_symbol = extract_ticker(company)
                        if ticker_symbol and ticker_symbol.lower() != "private":
                            comparable_companies.append(ticker_symbol)
                    except Exception as e:
                        print(f"Error processing company: {company}")
                        print(f"Error details: {e}")
                        continue
            else:
                print(f"\n{ticker} not found in columns!")


            comparable_metrics = []
            for metric in final_metrics:
                considered_companies = 0
                metric_sum = 0
                for tick in comparable_companies:
                    if tick == 'PRIVATE' or tick=='Private' or tick.lower() == 'private':
                        metric_value = 0
                    else :
                        if metric == "EV/EBITDA":
                            metric_value = get_ev_ebitda(tick) 
                        elif metric == "P/E TTM":
                            metric_value = get_trailing_pe(tick)
                        elif metric == "P/NAV":
                            metric_value = get_price_book(tick)
                        elif metric == "Dividend Yield":
                            metric_value = get_divident_yield(tick)
                        elif metric == "P/S TTM":
                            metric_value = get_price_sales(tick)
                        elif metric == "P/AUM":
                            metric_value = get_assets_under_management_ratio(tick)
                        else:
                            metric_value = 0
    
                    if metric_value > 0 :
                        considered_companies += 1
                        metric_sum += metric_value

                comparable_metrics.append(metric_sum/considered_companies)


            comparable_debt_equity = []
            comparable_current_ratio = []
            comparable_upside = []
            comparable_companies_selected = []
            for tick in comparable_companies:
                try:
                    if(tick.lower() == "private") :
                        comparable_debt_equity.append(0)
                        comparable_current_ratio.append(0)
                        comparable_upside.append(0)
                    else:
                        comparable_info = yf.Ticker(tick).info
                        comparable_debt_equity.append(comparable_info['debtToEquity'])
                        comparable_current_ratio.append(comparable_info['currentRatio'])
                        current_price = comparable_info.get('currentPrice', comparable_info.get('regularMarketPrice'))
                        upside = ((comparable_info['targetMeanPrice'] / current_price) - 1) * 100
                        comparable_upside.append(upside)
                        comparable_companies_selected.append(tick)
                except:
                    comparable_debt_equity.append(0)
                    comparable_current_ratio.append(0)
                    comparable_upside.append(0)

            comparable_de_mean = sum(comparable_debt_equity) / len(comparable_debt_equity)
            comparable_cr_mean = sum(comparable_current_ratio) / len(comparable_current_ratio)
            comparable_up_mean = sum(comparable_upside) / len(comparable_upside)

            all_content = {
                'info': stock.info,
                'income_stmt': stock.income_stmt,
                'balance_sheet': stock.balance_sheet,
                'news': stock.news
            }

            # Get historical data for the chart
            end_date = datetime.now()
            start_date = end_date - timedelta(days=365)
            all_content['history_data'] = stock.history(start=start_date, end=end_date)

            # # Create PDF report
            print("\nGenerating PDF report...")
            create_pdf_report(ticker, all_content, final_metrics, comparable_metrics, comparable_de_mean, comparable_cr_mean, comparable_up_mean)
         

        except Exception as e:
            print(f"Error processing request: {e}")
            print("Please try again with a valid ticker symbol.")

if __name__ == "__main__":
    print("Stock Financial Analysis Tool")
    print("Data source: Yahoo Finance")
    main()