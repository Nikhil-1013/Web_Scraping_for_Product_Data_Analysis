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
        # Open Chrome
        options = webdriver.ChromeOptions()
        options.add_argument('--ignore-certificate-errors')
        options.add_argument('--ignore-ssl-errors')
        driver = webdriver.Chrome(options=options)

        # Opening specific URL
        driver.get(Urls)
        driver.minimize_window()
        proName =  driver.find_element(By.CSS_SELECTOR, "#pdp_product_name").text
        a_rating =  driver.find_elements(By.XPATH, '//div[@class = "jm-rating-filled jm-rating-imgsize"]/*')
        a_rating = a_rating[-1].get_attribute('id')
        t_review =  driver.find_element(By.XPATH, '//span[@class = "review-count jm-body-m-bold jm-fc-primary-60 jm-pl-xs"]').text
        proID = driver.find_element(By.XPATH, '//div[@class = "jm-body-s-bold jm-mr-xxs"]').text.split(" ")[2]
        print(proName,a_rating, t_review, proID)
        return proName,a_rating, t_review, proID
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
            a,b,c,d = capture(Urls, output_text) 
            j = j-1
            output_text.insert(tk.END, "\n"+str(j)+" Remaining out of "+k+"\n")
            print("\nPro_Name: ",a,"\nAvg_Rating: ",b,"\nTot_Reviews: ",c,"\nPID: ",d,"\nLink: ",Urls,"\n")
            output_text.insert(tk.END, "\nPro_Name: "+a+"\nAvg_Rating: "+b+"\nTot_Reviews: "+c+"\nPID: "+d+"\nLink: "+Urls+"\n")
            output_text.update() # Update the GUI 
            data.append({'Product Link': Urls, 'Product Name': a, 'Product ID': d, 'Average Rating': b, 'Total Reviews': c})
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