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

# How do things work together? 

* <b>Introduction of each file </b>
  1. Everything in the Packages folder is necessary!
  2. `Year_Shee_fun.py` is the first function you will need to parse the document. It processes a source Excel file, splitting sheets into separate yearly Excel files grouped by month. (This is necessary to prevent the GPT 4-o model from hallucinating and processing outside its prompt length)
  3. `Agent.py` is the actual AI-Agent Structure that implemented meta-prompt engineering(an edited version of what is published by [OpenAI Meta Prompt](https://platform.openai.com/docs/guides/prompt-generation?context=text-out)) to use the ChatPromptTemplate and ChatOpenAI model to extract key information.
  4. `Combine.py` is the document that cleans the DataFrame by removing unwanted [''] structures from string columns. Then it combines JSON files in the 'result_dir' into their respective Excel files based on year and sheet names derived from the JSON filenames.
  
