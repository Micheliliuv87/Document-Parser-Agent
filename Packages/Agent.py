import openai
from langchain_openai import ChatOpenAI
from langchain_core.prompts import ChatPromptTemplate
import os

# Load API key and organization from environment variables
api_key = os.getenv("OPENAI_API_KEY")
organization = os.getenv("OPENAI_ORGANIZATION")


def create_extract_information_chain():
    """
    Creates a ChatPromptTemplate and connects it to the ChatOpenAI model.
    """
    # Define the prompt template
    prompt_template = ChatPromptTemplate.from_template(
        template="""
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
    )

    # Initialize the ChatOpenAI model
    llm = ChatOpenAI(
        temperature=0,
        model="gpt-4",  # Use a supported model
        api_key=api_key,
        organization=organization,
    )

    # Create a chain by linking the prompt template to the LLM
    llm_chain = prompt_template | llm
    return llm_chain


def extract_information(text):
    """
    Uses the ChatPromptTemplate and ChatOpenAI model to extract key information.
    """
    # Create the chain
    llm_chain = create_extract_information_chain()

    # Pass the text as a dictionary with the key matching the input variable in the template
    return llm_chain.invoke({"text": text})
