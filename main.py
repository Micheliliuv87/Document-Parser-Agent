


from Packages.Parser import process_documents
from Packages.Combine import combine_json_to_excel

if __name__ == "__main__":
    data_directory = "data/"         # Directory containing Excel files
    results_directory = "results/"   # Directory containing JSON results
    process_documents(data_dir=data_directory,result_dir=results_directory)
    combine_json_to_excel(results_directory)