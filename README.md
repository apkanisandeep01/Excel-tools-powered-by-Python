# 📊 Advanced Excel Tools (Powered by Python)

Tired of repetitive, manual tasks in Excel? This web app provides a suite of powerful tools to automate common spreadsheet operations like combining, splitting, and cleaning files—no formulas or macros required\!

It's a simple, fast, and secure way to get your data work done in seconds.



## ✨ Live Demo

### [🚀 Launch the Excel Multi-Tool\!](https://exceladvance-editor-python-ctuqvfojzjq6xmmgbcmxql.streamlit.app/)



-----

## 🛠️ Our Suite of Tools

This app bundles four powerful features into one easy-to-use interface.

### 📂 Combine Files

Merge multiple Excel or CSV files into a single, master spreadsheet. Perfect for consolidating monthly reports, sales data, or any separated datasets.

  * **Use Case:** You have sales reports for January, February, and March in separate files. This tool will combine them into one file for year-to-date analysis.

### ✂️ Split by Column

Got a large file with data for different categories (like departments, regions, or products)? This tool will automatically split it into separate, neatly organized Excel files for each category. You can even download them all in a single ZIP file\!

  * **Use Case:** You have a master employee list. This tool can create a separate Excel file for the "Sales," "Marketing," and "Engineering" departments.

### 🗑️ Drop Columns

Quickly clean up your files by removing unnecessary or sensitive columns. You can apply the same column removal to multiple files at once.

  * **Use Case:** Before sharing a customer list, you want to remove columns containing private information like phone numbers or addresses.

### 👁️ View Selected Columns

Focus on what matters. This tool lets you pick only the columns you want to see, creating a simplified view of your data that you can then preview and download.

  * **Use Case:** Your spreadsheet has 50 columns, but you only need to see `Name`, `Email`, and `Purchase Date`. This tool creates that view instantly.

-----

## 📖 How to Use

Using the app is as easy as 1-2-3:

1.  **Select a Tool:** Choose the task you want to perform from the tabs at the top (`Combine`, `Split`, etc.).
2.  **Upload Your Data:** Drag and drop your Excel or CSV file(s). The app will show you a preview.
3.  **Configure & Download:** Select your options (like the column to split by or columns to drop) and click the download button to get your new, processed file(s).

-----

## 💡 Technology

This application is built with modern, open-source tools:

  * **Python:** The core programming language.
  * **Streamlit:** For creating the interactive web application.
  * **Pandas:** For high-performance data manipulation.

-----

## 💻 Running Locally

If you'd like to run this project on your own computer:

1.  **Clone the repository:**

    ```bash
    git clone https://github.com/apkanisandeep01/Excel-tools-powered-by-Python.git
    cd Excel-tools-powered-by-Python
    ```

2.  **Install the necessary libraries:**

    ```bash
    pip install streamlit pandas openpyxl
    ```

3.  **Run the Streamlit app:**

    ```bash
    streamlit run app.py
    ```

    A new tab will open in your web browser with the app running locally.
