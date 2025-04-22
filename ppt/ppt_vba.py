import win32com.client
import os
import openai
from dotenv import load_dotenv
import time
import logging
import re # For parsing VBA code

from pptx import Presentation

CUR_PATH = os.path.dirname(__file__)

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

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def convert_pptx_to_pptm(pptx_path: str, pptm_path: str) -> bool:
    """
    Converts a .pptx file to a .pptm file using PowerPoint's SaveAs method.

    Args:
        pptx_path: Path to the source .pptx file.
        pptm_path: Path to save the destination .pptm file.

    Returns:
        True if conversion was successful, False otherwise.
    """
    if not os.path.exists(pptx_path):
        logging.error(f"Source file not found: {pptx_path}")
        return False

    if not pptx_path.lower().endswith(".pptx"):
        logging.warning(f"Source file is not a .pptx file: {pptx_path}")
        # You might still want to proceed if the user insists, but it's unusual.
        # return False # Or continue cautiously

    if not pptm_path.lower().endswith(".pptm"):
        logging.warning(f"Output path {pptm_path} should ideally end with .pptm. Adjusting.")
        pptm_path = os.path.splitext(pptm_path)[0] + ".pptm"

    powerpoint = None
    presentation = None
    success = False

    try:
        logging.info("Starting PowerPoint application...")
        powerpoint = win32com.client.DispatchEx("PowerPoint.Application")
        # Keep PowerPoint invisible during conversion
        # powerpoint.Visible = False

        logging.info(f"Opening source file: {pptx_path}")
        # Ensure paths are absolute for COM
        abs_pptx_path = os.path.abspath(pptx_path)
        abs_pptm_path = os.path.abspath(pptm_path)

        presentation = powerpoint.Presentations.Open(abs_pptx_path, WithWindow=False)

        logging.info(f"Saving as .pptm format to: {abs_pptm_path}")
        # FileFormat Enumeration for .pptm is 25 (ppSaveAsMacroEnabledPresentation)
        presentation.SaveAs(abs_pptm_path, FileFormat=25)
        success = True
        logging.info("File successfully saved in .pptm format.")

    except Exception as e:
        logging.error(f"Error during conversion: {e}", exc_info=True)
        success = False
    finally:
        # Ensure resources are released
        if presentation:
            try:
                presentation.Close()
            except Exception as e_close:
                logging.error(f"Error closing presentation: {e_close}")
        if powerpoint:
            try:
                powerpoint.Quit()
            except Exception as e_quit:
                logging.error(f"Error quitting PowerPoint: {e_quit}")
        # Clean up COM objects
        presentation = None
        powerpoint = None
        # import gc
        # gc.collect() # Force garbage collection if needed

    return success
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
    Now, generate new VBA code for Microsoft PowerPoint.
    This new code should perform a similar function to the template (e.g., creating slides with titles and bullet points), but it must attempt to extract relevant information from the following user-provided natural language data to create a PowerPoint presentation:

    {user_data}
    Instructions for Generation:

    Create a main public subroutine named {VBA_ENTRY_POINT_MACRO}. This subroutine will be called to generate the presentation content.
    Inside {VBA_ENTRY_POINT_MACRO}, intelligently parse the provided natural language data to identify potential slide titles and bullet points. This might involve looking for headings, lists, or sentences that could serve as titles or main points.
    For each identified potential slide:
    Add a new slide.
    Set the slide title using the extracted title information. If no clear title is found, use a generic title like "Slide X".
    Add any identified bullet points to the slide's content placeholder. If the template used a specific layout or placeholder index, try to replicate that. If not, use standard methods like Shapes.Placeholders(2).
    The generated code should be self-contained and runnable within a standard VBA module in PowerPoint.
    Do not include the original template code in your response unless it's being adapted.
    Focus on generating only the VBA code itself, without any introductory text, explanations, or markdown formatting around the code block.
    Make sure the generated VBA code is syntactically correct.
    Handle potential errors gracefully within the VBA if possible (e.g., check if placeholders exist).
    Use Option Explicit at the beginning of the module.
    Generated VBA Code:
    """

    try:
        # Using the chat completions endpoint (recommended)
        response = openai.chat.completions.create(
            model="gpt-4o-mini",  # Or "gpt-3.5-turbo", choose based on need/cost
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
    # template_file = os.path.join(CUR_PATH, "sample_ppt.pptx")
    # if template_file.endswith(".pptx"):
    #     # Convert to .pptm if necessary
    #     old_template_file = template_file
    #     template_file = old_template_file.replace(".pptx", ".pptm")
    #     convert_pptx_to_pptm(old_template_file, template_file)

    # Example: C:\\Users\\YourUser\\Documents\\Template.pptm
    output_file = "sample_output.pptm"
    # Example: C:\\Users\\YourUser\\Documents\\GeneratedPresentation.pptm

    # # Ensure absolute paths are used for COM interaction
    # template_file = os.path.abspath(template_file)
    # output_file = os.path.abspath(output_file)

    # # 1. Extract VBA from template
    # logging.info("--- Step 1: Extracting VBA ---")
    # template_vba_code = extract_vba_from_ppt(template_file)
    template_vba_code = """
Sub FillTextBoxesAutomatically()
    Dim sld As Slide
    Dim shp As Shape
    Dim contentList As Variant
    Dim i As Integer

    ' 채워넣을 텍스트 배열 (필요에 맞게 수정 가능)
    contentList = Array("김미리 프로필", "20YY.03.06", "0000@miridih.com", _
                        "미리코스메틱 신제품 네이밍 공모전", "미리대학교 홍보 모델", _
                        "퍼스널 브랜딩 스토어 운영")

    ' 첫 번째 슬라이드에 텍스트 채워넣기
    Set sld = ActivePresentation.Slides(1)

    i = 0
    For Each shp In sld.Shapes
        If shp.HasTextFrame Then
            If shp.TextFrame.HasText Then
                shp.TextFrame.TextRange.Text = contentList(i Mod UBound(contentList) + 1)
                i = i + 1
                If i > UBound(contentList) Then Exit For
            End If
        End If
    Next shp

    MsgBox "슬라이드 텍스트 자동 입력 완료!"
End Sub
"""

    if template_vba_code:
        logging.info("Template VBA extracted successfully.")
        # Log a snippet for verification (optional)
        # logging.debug(f"Template VBA (start):\n{template_vba_code[:300]}")

        # 2. Get user input
        logging.info("--- Step 2: Getting User Input ---")
        # user_content = get_user_input()
        user_content = """
        <div>
  <!--Header-->
  
  ![header](https://capsule-render.vercel.app/api?type=venom&color=gradient&height=300&section=header&text=Germanus'%20GitHub)
  
</div>

<div>
  <!--Body-->
  
  ## 👀 About Me
  #### :fire: AI / Backend / DevOps 개발자가 되기 위해 공부하고 있습니다.<br/>
  #### :mortar_board: 경상국립대학교(GNU), 항공우주및소프트웨어공학전공

  ### BOJ Rating
  [![Solved.ac 프로필](https://mazassumnida.wtf/api/v2/generate_badge?boj=qja1998)](https://solved.ac/qja1998)
  <br/>
  
  ## Main Experience

  ### **2021**
  - **[경상대 소프트웨어 구조 및 진화 연구실](https://www.gnu.ac.kr/soft/cm/cntnts/cntntsView.do?mi=13887&cntntsId=6492)**
    - [직책]
      - 학부 연구생
    - 관련 활동은 📚로 표시
  - **[BookCafe](https://saleese-gnu.github.io/bookcafe/)**
    - 카페 예약 시스템
    - 개발 인원: 4인
    - 개발 기간: 3개월
    - 역할: Andriod App 부분 개발(Kotlin)

  ### **2022**

  - **DIYA AI 연합 동아리**
    - [Dacon](https://dacon.io/myprofile/421883/home) 경진대회 참여
    - ~~[VAE 기반의 음성 style 변경 프로젝트](https://github.com/qja1998/audio)~~
  - **[경상대 SW 개발론 페이지 개발](https://saleese-gnu.github.io/)** 📚
    - Ruby 기반 GitHub page 구현 (1인 개발)
  - **코딩 하루 학원 강사**
    - [Streamlit 기반의 style transfer 앱 특강](https://github.com/qja1998/style_transform_with_streamlit)

  ### **2023**

  - **[AI 기반 탄소 배출량 관리 시스템](https://github.com/qja1998/co2-emission-management)**
    - [개요]
      - 기업의 탄소 배출량을 추적, 예측, 분석하여 관리가 용이하도록 하는 시스템 개발
    - [직책/역할]
      - 팀장
      - Backend(Django)
      - AI(탄소 배출량 예측)
      - 환경 관리(Docker)
  - **[BERT 기반 LLM 연구 시작](https://github.com/qja1998/pretrain_issue_bert)** 📚
    - [개요]
      - SW Issue Report에 특화된 LLM 제시 및 분류 성능 개선
    - [역할]
      - 언어 모델 pre-training
  - **빅데이터 시스템 소프트웨어 연구실**
    - AI 기반 탄소 배출량 관리 시스템의 고도화 및 주요 기능 특허 출원
    - [직책]
      - 외부 인력(학부 연구생)
      - 탄소 배출량 예측 모델 개선
      - 특허 출원 기능 자문

  ### **2024**

  - **[A Comparison of Pretrained Models for Classifying Issue Reports, IEEE Access](https://ieeexplore.ieee.org/document/10546475)** 📚
    - BERT 기반 연구가 완료되어 게재
  - **경상국립대학교(GNU), 항공우주및소프트웨어공학전공 졸업**
  - SSAFY 12기 - DATA track 1기 1학기 수료
    - [알고리즘 스터디 진행](https://github.com/qja1998/SSAFY_algorithm_study) - 스터디장
    - [Docker 스터디 진행](https://github.com/qja1998/SSAFY-Docker-Study)
  - DPG 해커톤 본선(전국 10위 이내) 진출
    - [RAG 기반 금융 도우미 및 상품 추천 시스템 개발](https://github.com/qja1998/nunuDream_rag)

  ### **2025**

  - SSAFY 12기 - DATA track 1기 2학기 진행중
    - MoMoSo 개발
      - [개요]
        - AI 기반 소설 작성
        - 소설 실시간 음성 토론
      - [역할]
        - Infra(Docker, GitLab CI)
        - RAG(Langchain)
        - 이미지 생성(Stable Diffusion)
    - [알고리즘 스터디 진행](https://github.com/qja1998/CoyoTe) - 스터디장

  ### Achievement

  - **2023 캡스톤디자인 작품 전시 및 발표회** - 금상
  - **2023 스마트 시티&모빌리티 캡스톤디자인 경진대회** - 은상
  - **탄소 배출량 예측 및 관리 시스템, 그리고, 그 방법** - 특허 출원
  - **2023 우수성과발표회** - 우수상(개인)
  - **[A Comparison of Pretrained Models for Classifying Issue Reports, IEEE Access](https://ieeexplore.ieee.org/document/10546475)** 📚
  - **DPG 해커톤 본선**

  <br/>
  
  ## 🧱 Tech Stack
  ### Language
  <!--Python-->
  <img src="https://img.shields.io/badge/Python%20IDLE-3776AB?logo=python&logoColor=fff"/>
  <!--MySQL-->
  <img src="https://img.shields.io/badge/MySQL-4479A1?logo=mysql&logoColor=fff"/>
  <br/>
  
  ### AI / Data Science
  <!--PyTorch-->
  <img src="https://img.shields.io/badge/PyTorch-EE4C2C?&logo=PyTorch&logoColor=white"/>
  <!--Hugging Face-->
  <img src="https://img.shields.io/badge/Hugging%20Face-FFD21E?logo=huggingface&logoColor=000"/>
  <!--Pandas-->
  <img src="https://img.shields.io/badge/Pandas-150458?logo=pandas&logoColor=fff)"/>
  <!--Numpy-->
  <img src="https://img.shields.io/badge/NumPy-4DABCF?logo=numpy&logoColor=fff"/>
  <!--Matplotlib-->
  <img src="https://custom-icon-badges.demolab.com/badge/Matplotlib-71D291?logo=matplotlib&logoColor=fff"/>
  <br/>
  
  ### CI/CD
  <!--Docker-->
  <img src="https://img.shields.io/badge/docker-2496ED?&logo=docker&logoColor=white"/>
  <!--GitLab CI-->
  <img src="https://img.shields.io/badge/GitLab%20CI-FC6D26?logo=gitlab&logoColor=fff"/>
  <!--GitLab CI/CD-->
  <img src="https://img.shields.io/badge/Jenkins-D24939?logo=jenkins&logoColor=white"/>

  ### Backend
  <!--Django-->
  <img src="https://img.shields.io/badge/Django-092E20?&logo=Django&logoColor=white"/>
  <!--FastAPI-->
  <img src="https://img.shields.io/badge/FastAPI-009485.svg?logo=fastapi&logoColor=white"/>
  <br/>
  

  ### Tools
  <!--git-->
  <img src="https://img.shields.io/badge/git-F05032?&logo=git&logoColor=white"/>
  <!--github-->
  <img src="https://img.shields.io/badge/GitHub-%23121011.svg?logo=github&logoColor=white"/>
  <!--jupyter-->
  <img src="https://img.shields.io/badge/jupyter-F37626?&logo=jupyter&logoColor=white"/>
  <!--notion-->
  <img src="https://img.shields.io/badge/notion-000000?&logo=notion&logoColor=white"/>
  <!--colab-->
  <img src="https://img.shields.io/badge/Google%20Colab-F9AB00?logo=googlecolab&logoColor=fff"/>
  <br/>

  ### ETC.
  <!--Selenium-->
  <img src="https://img.shields.io/badge/Selenium-43B02A?logo=selenium&logoColor=fff"/>
  <!--Anaconda-->
  <img src="https://img.shields.io/badge/Anaconda-44A833?logo=anaconda&logoColor=fff"/>
  <br/>
  
  ## 🤔 Github Stats
  ![](https://github-profile-summary-cards.vercel.app/api/cards/profile-details?username=qja1998&theme=nord_dark)

  ![](https://github-profile-summary-cards.vercel.app/api/cards/repos-per-language?username=qja1998&theme=nord_dark)
  ![](https://github-profile-summary-cards.vercel.app/api/cards/most-commit-language?username=qja1998&theme=nord_dark)

  ![](https://github-profile-summary-cards.vercel.app/api/cards/stats?username=qja1998&theme=nord_dark)
  
  ## Contact

  - **Blog**
    <!--Blog-->
    <a href="https://qja1998.github.io/">
      <img src="https://img.shields.io/website-up-down-green-red/https/qja1998.github.io"/>
    </a>
  - **Mail**
    <!--Mail-->
    <a href="mailto:rnjsrljqa98@gmail.com">
      <img src="https://img.shields.io/badge/gmail-EA4335?&logo=gmail&logoColor=white"/>
    </a>
</div>

        """

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