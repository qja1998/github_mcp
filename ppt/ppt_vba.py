import win32com.client
import os
import openai
from dotenv import load_dotenv
import time
import logging
import re # For parsing VBA code

from pptx import Presentation

import aspose.slides as slides

# --- Configuration ---
load_dotenv()
openai.api_key = os.getenv("OPENAI_API_KEY")
if not openai.api_key:
    raise ValueError("OpenAI API key not found. Please set it in the .env file.")

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- Constants ---
# Name of the main subroutine in the *generated* VBA code that should be run
# This name should be consistent with what you instruct the AI to generate.
VBA_ENTRY_POINT_MACRO = "CreatePresentationFromData"

# --- Functions ---

def convert_pptx_to_pptm(pptx_path, pptm_path):
    """
    Converts a .pptx file to .pptm format.

    Args:
        pptx_path (str): Path to the input .pptx file.
        pptm_path (str): Path to save the output .pptm file.
    """
    print(f"Converting {pptx_path} to {pptm_path}...")

    pptm_path   

def extract_vba_from_ppt(ppt_path: str) -> str | None:
    """
    Extracts VBA code from the standard modules of a PowerPoint file.

    Args:
        ppt_path: Path to the .pptm file.

    Returns:
        A string containing the combined VBA code from all standard modules,
        or None if an error occurs or no code is found.
    """
    if not os.path.exists(ppt_path):
        logging.error(f"Template PPT file not found: {ppt_path}")
        return None

    powerpoint = None
    presentation = None
    vba_code = ""
    extracted = False

    try:
        logging.info(f"Attempting to open PowerPoint and template: {ppt_path}")
        # Using DispatchEx for potentially better process management
        powerpoint = win32com.client.DispatchEx("PowerPoint.Application")
        # Make PowerPoint invisible during extraction (optional)
        # powerpoint.Visible = False

        presentation = powerpoint.Presentations.Open(ppt_path, WithWindow=False) # Open hidden
        logging.info("Template PPT opened successfully.")

        # Check if the presentation has a VBA project
        if not presentation.HasVBProject:
            logging.warning(f"Template file {ppt_path} does not contain a VBA project.")
            return None

        vb_project = presentation.VBProject
        logging.info("Accessing VBA Project...")

        # Iterate through components (modules, classes, forms)
        # We are primarily interested in standard modules (Type=1)
        for component in vb_project.VBComponents:
            if component.Type == 1: # vbext_ct_StdModule
                module_name = component.Name
                code_module = component.CodeModule
                lines_count = code_module.CountOfLines
                if lines_count > 0:
                    code = code_module.Lines(1, lines_count)
                    vba_code += f"'--- Module: {module_name} ---\n"
                    vba_code += code + "\n\n"
                    extracted = True
                    logging.info(f"Extracted code from module: {module_name}")

        if not extracted:
             logging.warning("No VBA code found in standard modules.")
             return None

        return vba_code

    except Exception as e:
        logging.error(f"Error extracting VBA: {e}", exc_info=True)
        # Consider more specific error handling (e.g., for COM errors)
        return None
    finally:
        # Ensure PowerPoint resources are released
        if presentation:
            try:
                presentation.Close()
                logging.info("Closed template presentation.")
            except Exception as e_close:
                 logging.error(f"Error closing presentation: {e_close}")
        if powerpoint:
            try:
                # Only quit if we started this instance (tricky to guarantee with DispatchEx)
                # Check if other presentations are open before quitting
                # if powerpoint.Presentations.Count == 0:
                powerpoint.Quit()
                logging.info("Quit PowerPoint application instance used for extraction.")
                # Small delay to ensure process termination
                time.sleep(1)
            except Exception as e_quit:
                 logging.error(f"Error quitting PowerPoint: {e_quit}")
        # Optional: Force release COM objects if issues persist
        # presentation = None
        # powerpoint = None
        # import gc
        # gc.collect()

def get_user_input() -> dict:
    """
    Collects user input for the presentation content.
    (This is a simple example, customize based on your needs)
    """
    logging.info("Collecting user input...")
    content = {}
    content['title'] = input("Enter the main title for the presentation: ")
    content['slides'] = []
    while True:
        slide_title = input("Enter title for a new slide (or leave blank to finish): ")
        if not slide_title:
            break
        slide_points = []
        print("Enter bullet points for this slide (leave blank to finish):")
        while True:
            point = input("- ")
            if not point:
                break
            slide_points.append(point)
        content['slides'].append({'title': slide_title, 'points': slide_points})
    logging.info(f"User input collected: {content}")
    return content

def generate_vba_with_ai(template_vba: str, user_data: dict) -> str | None:
    """
    Uses OpenAI API to generate new VBA code based on a template and user data.

    Args:
        template_vba: The VBA code extracted from the template PPT.
        user_data: A dictionary containing the user's content.

    Returns:
        The AI-generated VBA code as a string, or None if an error occurs.
    """
    logging.info("Generating VBA code using OpenAI...")

    # --- IMPORTANT: Customize this prompt heavily based on your template's structure ---
    prompt = f"""
    Analyze the following template VBA code for PowerPoint:
    ```vb
    {template_vba}
    ```

    Now, generate **new** VBA code for Microsoft PowerPoint.
    This new code should perform a similar *function* to the template (e.g., creating slides with titles and bullet points), but it must use the following user-provided data:

    ```json
    {user_data}
    ```

    **Instructions for Generation:**
    1.  Create a main public subroutine named `{VBA_ENTRY_POINT_MACRO}`. This subroutine will be called to generate the presentation content.
    2.  Inside `{VBA_ENTRY_POINT_MACRO}`, use the provided JSON data to create the PowerPoint slides.
    3.  For each item in the 'slides' array in the JSON data:
        * Add a new slide.
        * Set the slide title using the 'title' field.
        * Add the bullet points from the 'points' array to the slide's content placeholder. If the template used a specific layout or placeholder index, try to replicate that. If not, use standard methods like `Shapes.Placeholders(2)`.
    4.  The generated code should be self-contained and runnable within a standard VBA module in PowerPoint.
    5.  Do **not** include the original template code in your response unless it's being adapted.
    6.  Focus on generating **only the VBA code** itself, without any introductory text, explanations, or markdown formatting around the code block.
    7.  Make sure the generated VBA code is syntactically correct.
    8.  Handle potential errors gracefully within the VBA if possible (e.g., check if placeholders exist).
    9.  Use `Option Explicit` at the beginning of the module.

    **Generated VBA Code:**
    """

    try:
        # Using the chat completions endpoint (recommended)
        response = openai.chat.completions.create(
            model="gpt-4-turbo",  # Or "gpt-3.5-turbo", choose based on need/cost
            messages=[
                {"role": "system", "content": "You are a helpful assistant that generates Microsoft PowerPoint VBA code."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.5, # Adjust creativity vs determinism
            max_tokens=2000  # Adjust based on expected code length
        )

        generated_code = response.choices[0].message.content.strip()

        # Basic cleaning: Remove potential markdown code fences
        generated_code = re.sub(r'^```vb\s*|\s*```$', '', generated_code, flags=re.MULTILINE)

        logging.info("AI generated VBA code successfully.")
        # Log first few lines for verification (be careful with sensitive data)
        logging.debug(f"Generated VBA (start):\n{generated_code[:200]}")
        return generated_code

    except Exception as e:
        logging.error(f"Error calling OpenAI API: {e}", exc_info=True)
        return None

def create_ppt_with_vba(generated_vba: str, output_path: str, entry_point_macro: str):
    """
    Creates a new PPT, injects VBA, runs it, and saves the presentation.

    Args:
        generated_vba: The AI-generated VBA code string.
        output_path: The path where the final .pptm file should be saved.
        entry_point_macro: The name of the Subroutine within the VBA to run.
    """
    powerpoint = None
    presentation = None
    success = False

    try:
        logging.info("Creating new PowerPoint instance for generation...")
        powerpoint = win32com.client.DispatchEx("PowerPoint.Application")
        # Make it visible to see the process (optional, good for debugging)
        powerpoint.Visible = True
        # Add a new presentation
        presentation = powerpoint.Presentations.Add()
        logging.info("New presentation created.")

        # Ensure it has a VBA project (adding a module usually does this)
        # Add a standard module to the VBA project
        logging.info("Adding VBA module...")
        vb_project = presentation.VBProject
        vb_module = vb_project.VBComponents.Add(1) # 1 = vbext_ct_StdModule
        module_name = "GeneratedAIModule"
        vb_module.Name = module_name
        logging.info(f"Module '{module_name}' added.")

        # Insert the generated VBA code
        logging.info("Injecting generated VBA code...")
        vb_module.CodeModule.AddFromString(generated_vba)
        logging.info("VBA code injected.")

        # Run the main macro specified
        full_macro_path = f"{module_name}.{entry_point_macro}"
        logging.info(f"Attempting to run macro: {full_macro_path}")

        # Give PowerPoint a moment to process the added code
        time.sleep(2)

        # --- Run the Macro ---
        # Use Application.Run. Be mindful of potential errors here.
        try:
             powerpoint.Run(full_macro_path)
             logging.info(f"Successfully executed macro: {full_macro_path}")
             success = True
        except Exception as e_run:
            logging.error(f"ERROR RUNNING MACRO '{full_macro_path}': {e_run}")
            logging.error("Potential issues: Macro name incorrect, syntax error in generated VBA, security settings blocking execution, or PowerPoint instability.")
            # Consider saving the PPT even if macro fails, for debugging VBA
            success = False # Mark as failed if macro execution errors


        # Save the presentation
        # Save as .pptm because it contains macros, even if they only ran once.
        # Or save as .pptx if you are SURE the VBA is only for generation
        # and doesn't need to persist. But safer to use .pptm.
        logging.info(f"Saving presentation to: {output_path}")
        # FileFormat Enumeration: ppSaveAsDefault (usually .pptx), ppSaveAsPresentation (.pptx),
        # ppSaveAsMacroEnabledPresentation (.pptm = 25)
        presentation.SaveAs(output_path, FileFormat=25) # 25 = ppSaveAsMacroEnabledPresentation
        logging.info("Presentation saved successfully.")

    except Exception as e:
        logging.error(f"Error during PowerPoint creation/automation: {e}", exc_info=True)
        success = False
    finally:
        # Close presentation (if open)
        if presentation:
            try:
                # Close without saving changes *again* if SaveAs was successful
                presentation.Close()
                logging.info("Closed the generated presentation.")
            except Exception as e_close:
                logging.error(f"Error closing generated presentation: {e_close}")
        # Quit PowerPoint application (if we started it)
        if powerpoint:
            try:
                # Check count before quitting? Might be complex with DispatchEx
                # if powerpoint.Presentations.Count == 0:
                powerpoint.Quit()
                logging.info("Quit PowerPoint application instance used for generation.")
                time.sleep(1) # Allow time for process exit
            except Exception as e_quit:
                logging.error(f"Error quitting PowerPoint: {e_quit}")
        # Clean up COM objects
        # presentation = None
        # powerpoint = None
        # import gc
        # gc.collect()

    return success

# --- Main Execution ---
if __name__ == "__main__":
    template_file = r"C:\Users\SSAFY\Desktop\repo\github_mcp\ppt\sample_ppt.pptx"
    if template_file.endswith(".pptx"):
        # Convert to .pptm if necessary
        old_template_file = template_file
        template_file = old_template_file.replace(".pptx", ".pptm")
        convert_pptx_to_pptm(old_template_file, template_file)

    # Example: C:\\Users\\YourUser\\Documents\\Template.pptm
    output_file = "sample_output.pptm"
    # Example: C:\\Users\\YourUser\\Documents\\GeneratedPresentation.pptm

    # Ensure absolute paths are used for COM interaction
    template_file = os.path.abspath(template_file)
    output_file = os.path.abspath(output_file)

    # 1. Extract VBA from template
    logging.info("--- Step 1: Extracting VBA ---")
    template_vba_code = extract_vba_from_ppt(template_file)

    if template_vba_code:
        logging.info("Template VBA extracted successfully.")
        # Log a snippet for verification (optional)
        # logging.debug(f"Template VBA (start):\n{template_vba_code[:300]}")

        # 2. Get user input
        logging.info("--- Step 2: Getting User Input ---")
        user_content = get_user_input()

        # 3. Generate new VBA using AI
        logging.info("--- Step 3: Generating VBA with AI ---")
        generated_vba_code = generate_vba_with_ai(template_vba_code, user_content)

        if generated_vba_code:
            # 4. Create PPT using the generated VBA
            logging.info("--- Step 4: Creating PowerPoint ---")
            creation_success = create_ppt_with_vba(generated_vba_code, output_file, VBA_ENTRY_POINT_MACRO)

            if creation_success:
                logging.info(f"--- Process Complete: Presentation saved to {output_file} ---")
            else:
                logging.error("--- Process Failed: Presentation could not be created or saved correctly. Check logs. ---")
                logging.warning(f"Generated VBA code was:\n{generated_vba_code}") # Log generated code if failed
        else:
            logging.error("Failed to generate VBA code using AI.")
    else:
        logging.error(f"Failed to extract VBA code from {template_file}.")

    logging.info("Script finished.")