# France Template Validation - VBA Excel File

## Description
This Excel template is a **validation-based tool** tailored for France-specific data requirements. Built using VBA, it automates data integrity checks and enforces strict compliance with business rules. The goal is to simplify data management for users who may not be proficient with advanced tools like CRM systems or master data software, while ensuring error-free submissions.

---

## Why This Template?  
The client expressed a key business challenge:  
- Many of their sales and technical staff are not proficient with CRM or master data tools.  
- A simplified Excel-based solution was preferred to handle validations directly in a familiar environment, reducing the need for frequent back-and-forth queries.  

This template addresses these needs by incorporating most validation rules directly into the Excel file, enabling efficient, error-free data handling.

---

## Key Features
### 1. **User-Friendly Design**:
- Works exclusively on the `Actual Import Template` sheet.
- Drop-down options for the `Status` column to eliminate manual input errors.
- Clear error messages and color-coded highlights guide users to correct issues quickly.

### 2. **Validation Rules**:
- **Sheet-Specific**:
  - The sheet name must remain `Actual Import Template`.
  - Insertion of rows/columns is not allowed.
  - Header names cannot be modified.
- **Mandatory Fields**:
  - Must not be blank or exceed specified character limits.
  - Missing or incorrect values trigger errors and highlight fields in **red**.

### 3. **Conditional Formatting**:
- **Status = Approved**:
  - All mandatory fields must be correctly filled.
  - Missing or invalid data is highlighted in **red**.
- **Status = Reject**:
  - Only `EAN` and `Status` should be populated.
  - Any `France Marketing Content` provided in this case will be flagged in **red**.

### 4. **Error Reporting**:
- On saving, all errors are summarized for user review.
- Errors must be resolved before the file can be saved successfully.

---

## Client Benefits
- **Reduced Dependency on CRM Tools**:
  - Eliminates the need for staff training on advanced tools.
  - Allows immediate validation within Excel, minimizing technical queries.
- **Efficient Error Resolution**:
  - Highlights issues directly within the file, saving time for both the client and the service provider.
- **Streamlined Communication**:
  - Reduces back-and-forth client queries for minor data corrections.

---

## Ease of Use
### Instructions for Users:
1. **Choosing Status**:
   - Use the drop-down menu for `Status` values (`Approved` or `Reject`) to prevent manual input errors.
2. **Handling EANs**:
   - Ensure all rows between the first and last `EAN` have valid entries.
   - Missing `EANs` will trigger an error message.
3. **Correcting Errors**:
   - If data in a row is invalid, delete the row instead of clearing individual cells.
4. **Saving the File**:
   - Address all highlighted errors before saving to avoid interruptions.

---

## Future Scope
This template is scalable and can evolve with business needs:
1. **Enhanced Rules**:
   - Incorporate additional business-specific validations as requirements grow.
2. **Localization**:
   - Adapt for use in other regions with country-specific rules.
3. **Data Integration**:
   - Enable API or database connections for automated data input/output.
4. **Advanced Reporting**:
   - Add summary reports for frequently occurring errors to improve training and data quality over time.

---

## Best Practices
1. Always use the provided drop-down options for the `Status` column.
2. Avoid adding or deleting columns, as it disrupts the VBA validations.
3. Keep `Marketing Content` fields blank for `Reject` status to prevent errors.
4. Ensure all mandatory fields are filled correctly before saving.

---

## Protection and Access
- **Protection**: The sheet is protected by default to maintain data integrity.  
- **Password**: `password`  
- Protection is reapplied automatically when the file is saved.

---

## How to Access and Use
1. Download the file from this repository.  
2. Open the file in Excel and start filling out data in the `Actual Import Template` sheet.  
3. Address any errors highlighted in the file before saving.  
4. Save the file once all validations are passed successfully.  

---

## Repository Details
- **File Name**: `France_Template_Validation.xlsm`  
- **Primary Language**: VBA for Excel  
- **Protected Password**: `password`  
- **Author**: Pratyush  

---

## Conclusion
This template is an easy-to-use, robust tool for managing France-specific data validation. Designed with the client’s business needs in mind, it bridges the gap between technical expertise and operational requirements, reducing dependency on complex tools while ensuring data accuracy.

For any questions or contributions, feel free to create an issue or submit a pull request.
