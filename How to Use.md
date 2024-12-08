# How to Use the France Template Validation - Simple Guide

This guide will walk you through how to use the **France Template Validation** Excel file step-by-step. The template has been designed to make data entry simple and error-free. Just follow the instructions below to get started.

---

## Steps to Use the Template

### 1. **Open the Template**  
- Download and open the Excel file in your system.
- File: [Marketing_B2C_France_Template_v3.0](https://github.com/mr-pratyush/RealTime_AdvancedExcel_Project/blob/main/Marketing%20B2C_France_Temlate_v3.0.xlsm)

### 2. **Fill in the Data**  
- Enter the required data in the `Actual Import Template` sheet.  
- Make sure you fill in all the necessary fields for each row. The **mandatory fields** are marked, and they cannot be left blank.

### 3. **Mandatory Field Rules**  
- **EAN/UPC/BARCODE FOR BOTTLE**: This is a **mandatory field** and must always be filled in, whether the `Status` is **Approved** or **Reject**.  
- **Status**: If the **EAN/UPC/BARCODE FOR BOTTLE** field is filled in, the **Status** field must also be selected (either **Approved** or **Reject**).  
- **Short Description_FR**: This is a mandatory field and should have no more than **35 characters**.  
- **Long Description_FR**: This is also a mandatory field and must not exceed **70 characters**.  
- **Marketing Description_FR**: This field is required but should not exceed **4000 characters**.

### 4. **Choose the Status**  
- In the `Status` column, select one of the following options:
  - **Approved**: If the data is ready to be accepted.
  - **Reject**: If the data is not valid and needs correction.

### 5. **Check for Errors**  
- If you have chosen **Approved**, the template will automatically check for any missing or incorrect data.
  - Any errors will be highlighted in **red**. Fix these issues before saving.
- If you have chosen **Reject**, make sure **all France-related columns** (such as `Short Description_FR`, `Long Description_FR`, `Marketing Description_FR`) are **empty**. If any of these columns are filled, they will be flagged in **red**.

### 6. **Saving the File**  
- Once all the errors are fixed, you can save the file.
- If there are still errors, the template will show a message letting you know what needs to be corrected before saving.

### 7. **Protection and Access**  
- The file is protected by default to ensure data integrity.  
- To make changes to the template or unprotect the sheet, you will need the password: `password`.  
- The sheet will be re-protected automatically once you save the file.

---

## Tips for Best Results
- Always choose the `Status` from the provided drop-down list.  
- Do not add or delete columns, as it will affect the validation rules.  
- Delete the entire row for rows with invalid data instead of clearing individual cells.

---

## Troubleshooting
- Review the highlighted red fields and correct the data as needed if you encounter any issues or error messages.
- For any missing mandatory fields or incorrect entries, the template will automatically flag them for you.
