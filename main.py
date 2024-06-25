import gradio as gr
import threading
import os
from pathlib import Path
# from utils import merge_excels_by_sheet_name, merge_and_save, get_min_max_daterange, get_min_max_date_string
from utils2 import merge_excels, find_headers
import webbrowser
# from utils import merge_excels_by_sheet_name, merge_and_save, get_min_max_daterange, get_min_max_date_string

def file_upload(files):
    return files

def change_file_name(output_text):
    return output_text

def upload_and_merge(files, output_path="Merged_Report.xlsx"):
    merge_excels(files, output_file=output_path)
    return [gr.Files(label="Download Merged Report", value=["Merged_Report.xlsx"], visible=True), gr.DownloadButton(label=f"Download {output_path}", value=output_path, visible=True)]
    
def download_files_fn(files):
    return files

with gr.Blocks() as demo:
    gr.Markdown("<h1 style='text-align: center;'>GST Merger</h1><h2 style='text-align: center;'>Upload Excel Files to Merge</h2>")
    files = gr.Files(label="Upload Documents and Medical Reports", type="filepath", file_types=["xlsx"])
    # upload_button = gr.UploadButton(label="Upload Documents and Merge Them", type="filepath", file_count='multiple', file_types=["pdf", "docx", "jpg", "jpeg", "png"], )
    # output_text = gr.Textbox(label="Output File Path", type="text", value="Merged_Report.xlsx", placeholder="Enter the output file path e.g. Merged_Report.xlsx")
    # output_text.change(change_file_name, output_text, output_text)
    upload_button = gr.Button("Upload Files and Merge")
    files.upload(file_upload, files, files)

    download_files = gr.Files(label="Download Merged Report",visible=True)
    download_button = gr.DownloadButton(label="Download Merged Report", visible=False, value="Merged_Report.xlsx")
    
    upload_button.click(upload_and_merge, files, [download_files, download_button])


if __name__ == "__main__":
    # Create a thread for launching the demo
    demo_thread = threading.Thread(target=demo.launch)

    # Start the demo thread
    demo_thread.start()

    # Open default browser after the application has started running
    webbrowser.open('http://localhost:7860')

    # Wait for the demo thread to finish
    demo_thread.join()