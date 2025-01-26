# DocumentParserAgent

Document Parser AI-Agent Using OpenAI ChatGPT 4-o Model

# What is this for? 

* To parse an Excel file containing product updates for Google’s business-focused bundles.
* The goal is to extract individual features from the data (e.g., API for Google Tasks), categorize them by whether they were added or removed, and map them to the affected product bundles (e.g., Google Apps for Business and Education).
* The extracted information will be organized into a table with features as rows and product bundles as columns, creating one worksheet per month.
* This process involves iterating through the data worksheet by worksheet and building monthly and yearly summaries systematically.

# Where does it implement LLM? 

* The original file structure is Similar to the following
`Google Apps Product Updates (Synthetic Data)`

| Date       | Title                                              | Features                                                                                                               | Editions                          |
|------------|----------------------------------------------------|-----------------------------------------------------------------------------------------------------------------------|-----------------------------------|
| 01/15/2023 | Collaborative Whiteboard Added to Google Apps      | A new whiteboard feature allows real-time collaboration across all editions of Google Apps.                           | Basic, Pro, Enterprise            |
| 02/10/2023 | Advanced Admin Panel Introduced                   | A redesigned admin panel offers granular control, improved security settings, and analytics for administrators.        | Pro, Enterprise                   |
| 03/05/2023 | Task Automation Through Google Assistant           | Users can now automate tasks like scheduling and reminders using integrated Google Assistant features.                 | Basic, Pro, Enterprise            |
| 04/20/2023 | Enhanced Spreadsheet Visualization Tools           | Google Sheets now supports advanced charting, including heatmaps and pivot chart creation.                             | Basic, Pro, Enterprise            |
| 05/15/2023 | Unified Email and Chat Platform                    | Gmail and Google Chat are now seamlessly integrated into a single unified platform for easier communication.           | Pro, Enterprise                   |

* The Agent allows LLM to go through all the text in the document try to extract content under each feature and analyze if any notes contain updates to certain applications or features of Google's product.

* Prompt Template:

```
"""
        Given the following product update description, extract key information for each individual feature mentioned in the 'Features' text, and output as a JSON array.

        For each feature, extract:
        - "Date"
        - "Feature Name"
        - "Action" (whether the feature was added, removed, or updated)
        - "Products Affected" (list of products affected)

        ### Instructions:
        1. **Feature Name**: Identify the name of the feature, prioritizing the 'Title' field over the 'Features' field if both are provided.
        2. **Action**: Determine whether the feature was 'added', 'removed', or 'updated'. This should align with language in the 'Title' and 'Features' fields.
        3. **Products Affected**:
            - If the 'Title' or 'Features' mentions a specific product (e.g., "Gmail", "Google Docs"), use that product.
            - If no specific product is mentioned, fall back to the values in the 'Editions' field.
            - Do not include generic products unless explicitly mentioned.

        Input:
        {text}

        Output format:
        [
          {{
            "Date": "YYYY-MM-DD",
            "Feature Name": "Name of the feature",
            "Action": "added" or "removed" or "updated",
            "Products Affected": ["Product1", "Product2", ...]
          }},
          ...
        ]

        Please ensure the output is a valid JSON array.
"""
```

# How do things work together? [Packages Folder]

* <b>Introduction of each file </b>
  1. Everything in the Packages folder is necessary!
  2. `Year_Shee_fun.py` is the first function you will need to parse the document. It processes a source Excel file, splitting sheets into separate yearly Excel files grouped by month. (This is necessary to prevent the GPT 4-o model from hallucinating and processing outside its prompt length)
  3. `Agent.py` is the actual AI-Agent Structure that implemented meta-prompt engineering(an edited version of what is published by [OpenAI Meta Prompt](https://platform.openai.com/docs/guides/prompt-generation?context=text-out)) to use the ChatPromptTemplate and ChatOpenAI model to extract key information.
  4. `Combine.py` is the document that cleans the DataFrame by removing unwanted [''] structures from string columns. Then it combines JSON files in the 'result_dir' into their respective Excel files based on year and sheet names derived from the JSON filenames.
  5. `Parser.py` is the actual parser function that processes Excel files from a specified directory, extracting data from sheets containing required columns (Date, Title, Features, Editions). For each row, it formats the data into text, uses an external function (extract_information) to parse the text, and saves the parsed results as JSON files in the output directory. It handles errors, skips invalid rows, and logs issues for debugging.

# How to make this work? 

1. First, you will need a document that matches the description above: [Click here to jump back to the original file structure](#where-does-it-implement-llm).
2. Then run `prepare.py` to parse the document into smaller chunks for later processing.
3. Then run `main.py` to fully execute the work.
4. You then will get a document that is similar to the following structure


* The processed document should be similar to the following, where each document has a certain year as its title and multiple sheets for different months `Product Updates (Synthetic Data)`

| Date       | Feature Name                                        | Action   | Products Affected                                   |
|------------|----------------------------------------------------|----------|----------------------------------------------------|
| 2023-03-10 | Real-time Collaboration in Docs                    | added    | Google Docs, Google Apps, Google Workspace         |
| 2023-03-10 | Improved Spam Protection in Gmail                  | updated  | Gmail, Google Workspace                           |
| 2023-03-15 | Set Retention Policies for Drive Folders           | added    | Google Drive, Google Workspace                    |
| 2023-03-15 | Export Presentations to PDF                        | added    | Google Slides, Google Workspace                   |
| 2023-03-15 | Deprecated Older Export Formats                    | removed  | Google Slides, Google Docs, Google Sheets         |
| 2023-03-22 | Enhanced User Profiles for Google Meet             | added    | Google Meet, Google Workspace                     |
| 2023-03-22 | Edit Spreadsheets Offline                          | added    | Google Sheets, Google Workspace                   |


# Updates 

* Updated `Convert_to_FeatureSpecific.py` document to translate the generated document to look like the following:
`Feature Updates Timeline (Synthetic Data)`

| Year | Feature A                | Feature B                | Feature C                   | Feature D                   | Feature E                     | Feature F                   | Feature G                    | Feature H                   | Feature I                   | Feature J                    | Feature K                   |
|------|--------------------------|--------------------------|-----------------------------|-----------------------------|-------------------------------|-----------------------------|------------------------------|-----------------------------|-----------------------------|------------------------------|-----------------------------|
| 2006 | added                   |                          |                             |                             |                               |                             |                              |                             |                             |                              |                             |
| 2007 |                          |                          |                             |                             |                               |                             |                              |                             |                             |                              |                             |
| 2008 |                          |                          |                             |                             |                               |                             |                              |                             |                             |                              |                             |
| 2009 |                          |                          |                             |                             |                               |                             |                              |                             |                             |                              |                             |
| 2010 |                          |                          |                             |                             |                               |                             |                              |                             |                             |                              |                             |
| 2011 |                          | added                   |                             | added                      |                               |                             |                              |                             |                             |                              |                             |
| 2012 |                          | updated                 |                             |                             | added                        |                             |                              |                             |                             |                              |                             |
| 2013 |                          |                          | updated                     | added                      |                               | added                      |                              |                             |                             |                              |                             |
| 2014 | updated                  |                          |                             | updated                    |                               |                             |                              |                             |                             | added                        |                             |
| 2015 |                          |                          |                             |                             |                               |                             |                              |                             |                             |                              |                             |
| 2016 |                          |                          |                             |                             |                               |                             |                              |                             |                             |                              |                             |
| 2017 | added                   |                          |                             | added                      |                               |                             |                              |                             |                             |                              |                             |
| 2018 |                          |                          |                             |                             | added                        | added                      |                              |                             |                             |                              |                             |
| 2019 |                          | updated                 |                             | added                      |                               |                             |                              |                             |                             |                              |                             |

