# extract_pdf_to_excel/main.py
import configparser, sys
from modules import core, custom, pdfhandler
from pathlib import Path

if __name__ == "__main__":
    # Save filename and date/time info for later.
    filename = __file__.split("\\")[-1]
    curr_dt = core.get_curr_dt().strftime("%d%b%YZ%H%M%S").upper()
    curr_dt_tz = core.get_curr_dt().strftime("%Y-%m-%d %H:%M:%S %Z")

    # Parse the command line arguments.
    args = core.parse_args()

    # Read in the config file.
    config = configparser.ConfigParser()
    config.read(args.config)
    v = args.verbose

    # Set some variables from the config file.
    pdf_path = config["paths"]["pdf"]
    ocr_model_path = config["paths"]["model"]
    ocr_language = config["other"]["language"]
    out_dir = config["paths"]["output"]
    col_names = config["structure"]["cols"].split()
    del_imgs = eval(config["other"]["delete_images"]) # convert to Boolean
    del_pdfs = eval(config["other"]["delete_split_pdfs"])
    split_size = int(config["other"]["split_size"])

    # Set some more variables using the ones from the config (file 
    # extension added later for file_out).
    img = f"{str(out_dir)}\\{Path(pdf_path).stem}_{curr_dt}"
    file_out = f"{str(out_dir)}\\{Path(pdf_path).stem}_{curr_dt}.xlsx"
    file_ext = file_out.split(".")[-1]
    cols = len(col_names)
    
    # Print some helper text if program is in verbose mode.
    core.vprint(v, 
                    f"------------------BEGIN------------------\n"
                    f"> Running {filename} at {curr_dt_tz}...\n"
                    f"-----------------------------------------\n"
                    f"> Configuration variables set:\n"
                    f">> pdf_path: {pdf_path}\n"
                    f">> ocr_model_path: {ocr_model_path}\n"
                    f">> ocr_language: {ocr_language}\n"
                    f">> out_dir: {out_dir}\n"
                    f">> img: {img}\n"
                    f">> file_out: {file_out}\n"
                    f">> col_names: {col_names}\n"
                    f">> cols: {cols}\n"
                    f">> del_imgs: {del_imgs}\n"
                    f">> del_pdfs: {del_pdfs}\n"
                    f">> split_size: {split_size}\n"
                    f"-----------------------------------------")

    # Instantiate SpirePdf object (has some built-in error handling).
    try:
        pdf = pdfhandler.SpirePdf(pdf_path, col_names)
    except ValueError as e:
        sys.exit(e)

    # Check the OCR Model path from the config file.
    core.vprint(v, f"> Checking OCR path {ocr_model_path} exists:")
    if core.check_dir(ocr_model_path):
        core.vprint(v, f">> {ocr_model_path} confirmed.")
    else:
        sys.exit(f">> [ERROR] OCR Model could not be found in "
                 f"{ocr_model_path} folder/directory.")
    core.vprint(v, f"-----------------------------------------")

    # Check the Output folder/directory path from the config file.
    core.vprint(v, f"> Checking Output path {out_dir} exists:")
    if core.check_dir(out_dir):
        core.vprint(v, f">> {out_dir} confirmed.")
    else:
        sys.exit(f">> [ERROR] Output folder/directory does not exist: "
                 f"{out_dir}.")
    core.vprint(v, f"-----------------------------------------")

    # Set variable with the number of pages in the PDF file.
    page_count = pdf.pdf_page_count()

    # Output some more helper messages before processing the PDF file.
    core.vprint(v, f"> pdfhandler.SpirePdf() object created for "
                f"PDF document at {pdf_path}.")
    core.vprint(v, f">> Variables from SpirePdf object:")
    core.vprint(v, f">>> PDF File Name: {pdf.name()}")
    core.vprint(v, f">>> PDF File Page Count: {page_count}")
    core.vprint(v, f"-----------------------------------------")

    # If the PDF is larger than split_size (i.e., 10) pages, split it
    # into multiple documents to ensure there are no issues.
    core.vprint(v, f"> Checking if we need to split the PDF...")
    split = True
    if page_count <= split_size:
        split = False
        core.vprint(v, f">> Not splitting the PDF...")
    elif page_count % split_size == 0:
        # Divisible by split_size, so we can split our PDF exactly into
        # split_size (i.e., 10) individual PDF documents.
        split_doc_num = page_count / 10
    else:
        split_doc_num = page_count // 10 + 1
    
    if split:
        core.vprint(v, f">> PDF will be split into {split_doc_num} separate "
                    f"files for smoother processing...")
        
    core.vprint(v, f"-----------------------------------------")
    
    # Perform the actual splitting of the original PDF (if needed).
    split_pdf_list = []
    if split:
        # Page count is greater than split_size so we need to split the
        # PDF into multiple files.
        core.vprint(v, f"> Splitting PDF into {split_doc_num} files...")
        for i in range(split_doc_num):
            # Generate a name for the new, split PDF to be created.
            tmp_file = f"{out_dir}\\pdf_page{i}.pdf"
            split_pdf_list.append(tmp_file)

            # Split should start at index 0, 10, 20, etc. (assuming a
            # split_size of 10).
            split_start = i * split_size

            if i < split_doc_num - 1:
                # This is not the final range we are splitting, just
                # add split_size-1 so we get the correct ending index.
                split_end = split_start + (split_size - 1)
            else:
                # This is the final range we are splitting; use the
                # final page index.
                split_end = page_count - 1
            
            # Call our function to actually split the PDF using the now
            # identified start and end range.
            pdfhandler.split_pdf_on_range(pdf.pdf, tmp_file, split_start,
                                           split_end)
            core.vprint(v, f">> {tmp_file} created for Pages {split_start+1} "
                        f"thru {split_end+1}...")
            core.vprint(v, f"                 ----------")

        core.vprint(v, f"-----------------------------------------")
    else:
        # PDF does not need to be split, so just use original file.
        split_pdf_list.append(pdf_path)

    pdf_pages_as_list = []
    total_page_count = 0
    core.vprint(v, f"> Beginning process to scan {pdf.name()}, "
                f"convert to Excel, and save as {file_out}")
    core.vprint(v, f"                 ----------")
    # Go through the newly split PDF docs to extract relevant text for
    # each one.
    for split_path in split_pdf_list:
        # Load a new PdfDocument object for the split PDF.
        try:
            split_pdf = pdfhandler.SpirePdf(split_path, col_names)
        except ValueError as e:
            sys.exit(e)

        # Setup a custom name for the split PDF images to be created.
        tmp_split_name = f"{img}_{Path(split_path).stem}"

        # Cycle thru each PDF page and extract the text we want.
        for page_index in range(split_pdf.pdf_page_count()):
            core.vprint(v, f">> Starting extraction process for "
                        f"Page {page_index + 1} of {split_pdf.name()}:")

            # Setup the OCR scanner.
            core.vprint(v, f">>> Setting up OCR Scanner object...")
            scanner = split_pdf.scanner_init(ocr_language, ocr_model_path)
            
            # Save the current PDF page as an image.
            core.vprint(v, f">>> Saving Page {page_index + 1} as "
                        f"an Image...")
            tmp_img_path = f"{tmp_split_name}{page_index + 1}.png"
            split_pdf.save_as_img(page_index, tmp_img_path)

            # Scan the newly saved image of the current PDF page and
            # extract all the text as a string.
            core.vprint(v, f">>> Extracting text from Page "
                        f"{page_index + 1} Image...")
            recognized_text = split_pdf.scanner_to_text(scanner, tmp_img_path)

            # Split the string into a python list for easier handling.
            img_text_list = split_pdf.split_scanned_text(
                                        recognized_text, "\n")

            # Find the column header with the highest index in the list
            # of PDF text so we know where the real data starts.
            max_col = split_pdf.max_column_header(img_text_list)

            # Iterate through the lines of text from the image and add
            # them to pdf_pages_as_list to populate with text from each
            # page of the PDF.
            for i, line in enumerate(img_text_list):
                # Only want data between headers & last line of text 
                # (warning message).
                if i > max_col and i < len(img_text_list) - 1:
                    pdf_pages_as_list.append(line)
        
            # Delete the image we created if the del_imgs config
            # variable is set to True (now that we're done extracting
            # text from it).
            if del_imgs:
                core.vprint(v, f">>> Deleting image created to "
                            f"extract text from PDF Page {page_index + 1}...")
                msg = core.delete_file(tmp_img_path)
                core.vprint(v, f">>>> {msg}")

            total_page_count += 1
            core.vprint(v, f">>> Finished processing Page "
                        f"{page_index + 1}...")
            core.vprint(v, f"                 ----------")

        # Delete the split PDF file that was created if the requisite
        # config variable is True.
        if del_pdfs and split:
            core.vprint(v, f"                 ----------")
            core.vprint(v, f">> Deleting split PDF document...")
            msg = core.delete_file(split_pdf.filepath)
            core.vprint(v, f">>> {msg}")
            core.vprint(v, f"                 ----------")
        
        # Close the split PDF document object now that we are done.
        split_pdf.close()

    # Print helper message if we are in verbose mode and have finished
    # processing the PDF pages.
    pdf.close()
    core.vprint(v, f">> Finished extracting data from {pdf.name()}"
                         f"...converting data and saving to Excel now...")
    core.vprint(v, f"-----------------------------------------")

    # Instantiate a dataframe to store the PDF data (structure defined
    # via config file) and slice pdf_pages_as_list to populate the
    # columns of the dataframe.
    # breakpoint()
    pdf_dataframe = core.slice_list_to_df(pdf_pages_as_list, col_names)

    # Export the data to an Excel file.
    pdf_dataframe.to_excel(file_out)

    # Check if output file was successfully created and print out final
    # helper messages if in verbose mode.
    if core.check_file(file_out, file_ext):
        core.vprint(v, f"> Successfully converted PDF and saved as"
                    f" {file_out}.")
    else:
        sys.exit(f"> [ERROR] Something prevented the converted data from "
                 f"being saved to {file_out}.")
    

    core.vprint(v, f"-----------------RESULTS-----------------")
    core.vprint(v, f"> {filename} Completed at "
                f"{core.get_curr_dt().strftime("%Y-%m-%d %H:%M:%S %Z")}.")
    core.vprint(v, f"> Pages Converted: {total_page_count}")
    core.vprint(v, f"> Columns of Data: {pdf_dataframe.shape[1]}")
    core.vprint(v, f"> Rows of Data: {pdf_dataframe.shape[0]}")
    core.vprint(v, "-------------------END-------------------")
