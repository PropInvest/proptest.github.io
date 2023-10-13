import os
from flask import Flask, render_template, request, redirect, url_for, jsonify, send_file
from werkzeug.utils import secure_filename
from selenium import webdriver
from selenium.webdriver.chrome.options import Options as ChromeOptions
import pandas as pd
import re
from flask import send_file
from waitress import serve

cleaned_telephone1 = []
name1 = []
street1 = []
city1 = []
state1 = []

application = Flask(__name__)


UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xls', 'xlsx'}

# Ensure the 'uploads' directory exists
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

application.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
application.config['ALLOWED_EXTENSIONS'] = ALLOWED_EXTENSIONS

# Function to check if the uploaded file has a valid extension
def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in application.config['ALLOWED_EXTENSIONS']


def scrape_data(street, city, state):

    global cleaned_telephone1
    global name1
    street = street.lower().title()
    city = city.lower().title()
    state = state.lower().title()

    street1.append(street)
    city1.append(city)
    state1.append(state)

    url = f"https://www.usa-people-search.com/address/{street}/{city}-{state}"

    chrome_options = ChromeOptions()
    chrome_options.add_argument("start-maximized")

    with webdriver.Chrome(options=chrome_options) as driver:
        driver.get("view-source:" + url)
        html_content = driver.page_source

        start_index = html_content.find('"@type": "Person",')
        if start_index != -1:
            extracted_content = html_content[start_index:]
            link_pattern = re.compile(r'https?://\S+')
            first_link_match = re.search(link_pattern, extracted_content)

            if first_link_match:
                first_link = first_link_match.group()
                modified_link = first_link.strip('</td></tr><tr><td').replace('"', '')
                modified_link_with_view_source = "view-source:" + modified_link
                driver.get(modified_link_with_view_source)

                telephone_start_index = driver.page_source.find('"telephone": ["')
                if telephone_start_index != -1:
                    telephone_content = driver.page_source[telephone_start_index:]
                    comma_index = telephone_content.find(',')
                    if comma_index != -1:
                        extracted_content = telephone_content[:comma_index]
                        opening_parenthesis_index = extracted_content.find('(')
                        if opening_parenthesis_index != -1:
                            cleaned_telephone = extracted_content[opening_parenthesis_index:]
                            cleaned_telephone = cleaned_telephone.split('"', 1)[0]

                            prefix = "view-source:https://www.usa-people-search.com/"
                            modified_link_without_prefix = modified_link_with_view_source.replace(prefix, '')
                            modified_link_without_prefix = modified_link_without_prefix.split('/', 1)[0].replace('-',
                                                                                                                 ' ').title()
                            cleaned_telephone = cleaned_telephone
                            print("Cleaned Name:", modified_link_without_prefix)
                            cleaned_telephone1.append(cleaned_telephone)
                            print("Cleaned Telephone:", cleaned_telephone1)
                            name1.append(modified_link_without_prefix)

                            return cleaned_telephone1, name1



                        else:
                            return "Opening parenthesis '(' not found in the telephone content.", None
                    else:
                        return "No comma found after 'telephone':", None
                else:
                    return "Starting point 'telephone' not found in the HTML content.", None
            else:
                return "No valid link found after '@type': 'Person',", None
        else:
            return "Starting point '@type': 'Person' not found in the HTML content.", None


@application.route('/', methods=['GET', 'POST'])



def index():

    processed_data = []
    if request.method == 'POST':
        if 'file' not in request.files:
            return jsonify({'error': 'No file part'})

        file = request.files['file']
        if file.filename == '' or not allowed_file(file.filename):
            return jsonify({'error': 'Invalid file'})

        filename = secure_filename(file.filename)
        file.save(os.path.join(application.config['UPLOAD_FOLDER'], filename))

        # Read the uploaded Excel file
        excel_data = pd.read_excel(os.path.join(application.config['UPLOAD_FOLDER'], filename))

        # Check if either "Street Address" or "Address" column exists in the Excel file
        if 'Street Address' in excel_data.columns:
            address_column = 'Street Address'
        elif 'Address' in excel_data.columns:
            address_column = 'Address'
        else:
            return jsonify({'error': 'Required address column not found in the Excel file'})

        for index, row in excel_data.iterrows():
            global cleaned_telephone1
            global cleaned_telephone
            global name1
            global modified_link_without_prefix
            street = row[address_column].replace(' ', '-')
            city = row['City'].replace(' ', '-')
            state = row['State'].replace(' ', '-')
            cleaned_telephone, modified_link_without_prefix = scrape_data(street, city, state)

            if cleaned_telephone1 and modified_link_without_prefix:
                processed_data.append({

                    'cleanedTelephone': cleaned_telephone1,
                    'modifiedLink': name1
                })

        return render_template('index.html', data=processed_data, cleaned_telephone1=cleaned_telephone1, name1=name1, street1=street1, city1=city1, state1=state1)

    print(cleaned_telephone1)
    return render_template('index.html', data=processed_data,cleaned_telephone1=cleaned_telephone1, name1=name1, street1=street1, city1=city1, state1=state1)


@application.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'})

    file = request.files['file']

    if file.filename == '' or not allowed_file(file.filename):
        return jsonify({'error': 'Invalid file'})

    filename = os.path.join(application.config['UPLOAD_FOLDER'], file.filename)
    file.save(filename)

    # Handle data processing logic here
    try:
        # Read the uploaded Excel file
        excel_data = pd.read_excel(filename)

        # Process excel_data and update cleaned_telephone1 list
        # For example:
        # cleaned_telephone1 = excel_data['Telephone'].tolist()

        # Emit an update to connected clients when data is ready

        return jsonify({'status': 'success'})
    except Exception as e:
        return jsonify({'error': str(e)})


@application.route('/download')
def download_file():
    processed_data = pd.DataFrame({
        'Cleaned Telephone': cleaned_telephone1,
        'Cleaned Name': name1,
        'Cleaned Address': street1,
        'Cleaned City': city1,
        'Cleaned State': state1

    })
    # Save the DataFrame to an Excel file
    processed_data.to_excel('processed_data.xlsx', index=False)

    # Send the Excel file as a downloadable attachment
    return send_file('processed_data.xlsx', as_attachment=True)


if __name__ == '__main__':
    serve(application, host='0.0.0.0', port=5111)
