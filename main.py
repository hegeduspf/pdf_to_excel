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

    # Set some variables from the config file.
    pdf_path = config["paths"]["pdf"]
    ocr_model_path = config["paths"]["model"]
    ocr_language = config["other"]["language"]
    out_dir = config["paths"]["output"]
    col_names = config["structure"]["cols"].split()
    del_imgs = eval(config["other"]["delete_images"]) # convert to Boolean

    # Set some more variables using the ones from the config (file 
    # extension added later for file_out).
    img = f"{str(out_dir)}\\{Path(pdf_path).stem}_{curr_dt}"
    file_out = f"{str(out_dir)}\\{Path(pdf_path).stem}_{curr_dt}.xlsx"
    file_ext = file_out.split(".")[-1]
    cols = len(col_names)
    
    # Print some helper text if program is in verbose mode.
    core.vprint(args.verbose, 
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
                    f"-----------------------------------------")

    # Instantiate SpirePdf object (has some built-in error handling).
    try:
        pdf = pdfhandler.SpirePdf(pdf_path, col_names)
    except ValueError as e:
        sys.exit(e)

    # Check the OCR Model path from the config file.
    core.vprint(args.verbose, f"> Checking OCR path {ocr_model_path} exists:")
    if core.check_dir(ocr_model_path):
        core.vprint(args.verbose, f">> {ocr_model_path} confirmed.")
    else:
        sys.exit(f">> [ERROR] OCR Model could not be found in "
                 f"{ocr_model_path} folder/directory.")
    core.vprint(args.verbose, f"-----------------------------------------")

    # Check the Output folder/directory path from the config file.
    core.vprint(args.verbose, f"> Checking Output path {out_dir} exists:")
    if core.check_dir(out_dir):
        core.vprint(args.verbose, f">> {out_dir} confirmed.")
    else:
        sys.exit(f">> [ERROR] Output folder/directory does not exist: "
                 f"{out_dir}.")
    core.vprint(args.verbose, f"-----------------------------------------")

    # Set variable with the number of pages in the PDF file.
    page_count = pdf.pdf_page_count()

    # Output some more helper messages before processing the PDF file.
    core.vprint(args.verbose, f"> pdfhandler.SpirePdf() object created for "
                f"PDF document at {pdf_path}.")
    core.vprint(args.verbose, f">> Variables from SpirePdf object:")
    core.vprint(args.verbose, f">>> PDF File Name: {pdf.name()}")
    core.vprint(args.verbose, f">>> PDF File Page Count: {page_count}")
    core.vprint(args.verbose, f"-----------------------------------------")

    # Cycle thru each PDF page and extract the text we want.
    core.vprint(args.verbose, f"> Beginning process to scan {pdf.name()}, "
                f"convert to Excel, and save as {file_out}")
    core.vprint(args.verbose, f"                 ----------")
    pdf_pages_as_list = []
    for page_index in range(page_count):
        core.vprint(args.verbose, f">> Starting extraction process for Page "
                    f"{page_index + 1}:")
        
        # Setup the OCR scanner.
        core.vprint(args.verbose, f">>> Setting up OCR Scanner object...")
        scanner = pdf.scanner_init(ocr_language, ocr_model_path)

        # Save the current PDF page as an image.
        core.vprint(args.verbose, f">>> Saving Page {page_index + 1} as an "
                    f"Image...")
        tmp_img_path = f"{img}_{page_index + 1}.png"
        pdf.save_as_img(page_index, tmp_img_path)

        # Scan the newly saved image of the current PDF page and
        # extract all the text as a string.
        core.vprint(args.verbose, f">>> Extracting text from Page "
                    f"{page_index + 1} Image...")
        recognized_text = pdf.scanner_to_text(scanner, tmp_img_path)

        # Split the string into a python list for easier handling.
        img_text_list = pdf.split_scanned_text(recognized_text, "\n")

        # Find the column header with the highest index in the list of
        # PDF text so we know where the real data starts.
        max_col = pdf.max_column_header(img_text_list)

        # Iterate through the lines of text from the image and add
        # them to pdf_pages_as_list to populate with text from each
        # page of the PDF.
        for i, line in enumerate(img_text_list):
            # Only want data between headers & last line of text 
            # (warning message).
            if i > max_col and i < len(img_text_list) - 1:
                pdf_pages_as_list.append(line)
        
        # Delete the image we created if the del_imgs config variable
        # is set to True (now that we're done extracting text from it).
        if del_imgs:
            core.vprint(args.verbose, f">>> Deleting image created to extract "
                        f"text from PDF Page {page_index + 1}...")
            msg = core.delete_file(tmp_img_path)
            core.vprint(args.verbose, f">>>> {msg}")
        
        core.vprint(args.verbose, f">>> Finished processing Page "
                    f"{page_index + 1}...")
        core.vprint(args.verbose, f"                 ----------")

    # Print helper message if we are in verbose mode and have finished
    # processing the PDF pages.
    core.vprint(args.verbose, f">> Finished extracting data from {pdf.name()}"
                         f"...converting data and saving to Excel now...")
    core.vprint(args.verbose, f"-----------------------------------------")

    # Instantiate a dataframe to store the PDF data (structure defined
    # via config file) and slice pdf_pages_as_list to populate the
    # columns of the dataframe.
    pdf_dataframe = core.slice_list_to_df(pdf_pages_as_list, col_names)

    # Export the data to an Excel file.
    pdf_dataframe.to_excel(file_out)

    # Check if output file was successfully created and print out final
    # helper messages if in verbose mode.
    if core.check_file(file_out, file_ext):
        core.vprint(args.verbose, f"> Successfully converted PDF and saved as"
                    f" {file_out}.")
    else:
        sys.exit(f"> [ERROR] Something prevented the converted data from "
                 f"being saved to {file_out}.")
    

    core.vprint(args.verbose, f"-----------------RESULTS-----------------")
    core.vprint(args.verbose, f"> {filename} Completed at "
                f"{core.get_curr_dt().strftime("%Y-%m-%d %H:%M:%S %Z")}.")
    core.vprint(args.verbose, f"> Pages Converted: {page_index + 1}")
    core.vprint(args.verbose, f"> Columns of Data: {pdf_dataframe.shape[1]}")
    core.vprint(args.verbose, f"> Rows of Data: {pdf_dataframe.shape[0]}")
    core.vprint(args.verbose, "-------------------END-------------------")