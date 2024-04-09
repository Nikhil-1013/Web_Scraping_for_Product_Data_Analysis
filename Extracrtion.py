# Libraries 
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
import tkinter as tk
from tkinter import scrolledtext, filedialog
# import time

# Function to Capture text from a dynamic page
def capture(Urls, output_text):

    try:
        extracted_data = {}
        # Open Chrome
        options = webdriver.ChromeOptions()
        options.add_argument('--ignore-certificate-errors')
        options.add_argument('--ignore-ssl-errors')
        driver = webdriver.Chrome(options=options)

        # Opening specific URL
        driver.get(Urls)
        driver.minimize_window()
        extracted_data["Product Name"] =  driver.find_element(By.CSS_SELECTOR, "#pdp_product_name").text
        a_rating =  driver.find_elements(By.XPATH, '//div[@class = "jm-rating-filled jm-rating-imgsize"]/*')
        extracted_data["Avg Rating"] = a_rating[-1].get_attribute('id')
        extracted_data["Total Ratings"] =  driver.find_element(By.XPATH, '//span[@class = "review-count jm-body-m-bold jm-fc-primary-60 jm-pl-xs"]').text
        extracted_data["Product ID"] = driver.find_element(By.XPATH, '//div[@class = "jm-body-s-bold jm-mr-xxs"]').text.split(" ")[2]
        # print(proName,a_rating, t_review, proID)
        return extracted_data
    except Exception as e:
        output_text.insert(tk.END, f"Error occurred while processing this URL : \n{Urls}\n")
        output_text.update()# Update the GUI 
        print(f"Error occurred while processing {Urls}: {e}\n")
        return None, None, None, None
    finally:
        driver.stop_client()
        driver.close()

# Reading Excel file 
def process_urls(output_text):
    try:
        file_path = filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel Files", "*.xlsx"), ("All files", "*.*")])
        if not file_path:
            return
        df = pd.read_excel(file_path)
        j = len(df['Url'])
        k=j
        k=str(k)
        data  = []
        # Looping for Multiple links
        for Urls in df['Url']:
            # Calling capture function
            extracted_data = capture(Urls, output_text) 
            j = j-1
            output_text.insert(tk.END, "\n"+str(j)+" Remaining out of "+k+"\n")
            print("\nPro_Name: ",extracted_data["Product Name"],"\nAvg_Rating: ",extracted_data["Avg Rating"],"\nTot_Reviews: ",extracted_data["Total Ratings"],"\nPID: ",extracted_data["Product ID"],"\nLink: ",Urls,"\n")
            output_text.insert(tk.END, "\nPro_Name: "+extracted_data["Product Name"]+"\nAvg_Rating: "+extracted_data["Avg Rating"]+"\nTot_Reviews: "+extracted_data["Total Ratings"]+"\nPID: "+extracted_data["Product ID"]+"\nLink: "+Urls+"\n")
            output_text.update() # Update the GUI 
            data.append(extracted_data)
            default = data

        # Saving to Excel file
        if data:
            data = pd.DataFrame(data)
            # Asking where to save
            output_file_path = filedialog.asksaveasfilename(title="Save as", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
            print(data)
            if output_file_path:
                data.to_excel(output_file_path, "Ratings", index=False)
                print('Extracted data is available in ',output_file_path.split("/")[-1],'\n')
                output_text.insert(tk.END, 'Extracted data is available in '+output_file_path.split("/")[-1]+'\n')
                output_text.update()  # Update the GUI
    except Exception as e:
        print(e)
    finally:
        # Saving to Default Excel file
        if default:
            default = pd.DataFrame(default)
            default.to_excel("Default.xlsx", "Ratings", index=False)
            print("Copy of current extracted data is available in 'Default.xlsx' .\n")
            output_text.insert(tk.END, "\nCopy of current extracted data is available in 'Default.xlsx' .\n")
            output_text.update()  # Update the GUI

# Created Tkinter window
root = tk.Tk()
root.title("Product Data Extractor")

# Title label
title_label = tk.Label(root, text="Product Data Extractor", font=("Helvetica", 16, "bold"))
title_label.pack(pady=10)

# Created a text widget for displaying Output
output_text = scrolledtext.ScrolledText(root, width=45, height=20, font = ("Times New Roman", 15))
output_text.pack(padx=10, pady=10)

# Created a button to select the Excel file containing URLs
select_file_button = tk.Button(root, text="Select Excel File", command=lambda: process_urls(output_text))
select_file_button.pack(pady=10)

root.mainloop()