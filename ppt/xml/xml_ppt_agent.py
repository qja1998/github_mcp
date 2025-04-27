import asyncio
import shutil

import os
import subprocess
import dotenv
dotenv.load_dotenv(override=True)

from typing import List
from collections import deque
from xml.etree import ElementTree as ET

GITHUB_PERSONAL_ACCESS_TOKEN = os.getenv("GITHUB_PERSONAL_ACCESS_TOKEN")
OPENAI_MODEL = os.getenv("OPENAI_MODEL")
print(OPENAI_MODEL)

from agents import Agent, Runner, trace, WebSearchTool
from agents.mcp import MCPServer, MCPServerStdio

import json

CUR_PATH = os.getcwd()

TEMPLATE_XML_PROJECT_PATH = "C:/Users/kwon/Desktop/repo/github_mcp/ppt/xml/xml_template/ppt/slides/slide7.xml"
RESULT_XML_PROJECT_PATH = "C:/Users/kwon/Desktop/repo/github_mcp/ppt/xml//xml_result/ppt/slides/slide{slide_idx}.xml"
async def run(mcp_servers: List[MCPServer], content: str):
    print("Running...")
    agent = Agent(
        model=OPENAI_MODEL,
        name="Assistant",
        instructions=f"Answer questions about json",
        mcp_servers=mcp_servers,
        # tools=[WebSearchTool()]
    )

    print("Agent initialized.")

    PROMPT = f"""
    당신은 xml, pptx 전문가입니다. 다음 xml template는 pptx-compose의 xml template입니다.
    이 template을 바탕으로 사용자가 원하는 pptx를 생성하기 위한 xml을 생성해 주세요.
    """

    # 내용 뽑아내기
    extract_prompt = f"""
    다음 content에서 pptx를 생성하기 위한 프로젝트를 추출합니다.
    항상 오직 프로젝트 이름만을 **,**로 구분하여 답변합니다. **이외의 미사어구는 절대 붙이지 않습니다.**
    content: {content}
    project name:
    """
    result = await Runner.run(starting_agent=agent, input=extract_prompt, max_turns=30)
    projects = result.final_output.split(",")
    print("추출한 프로젝트:", result.final_output)
    # 내용 기반으로 슬라이드 개수만큼 반복
    for i, project in enumerate(projects, start=5):
        # 슬라이드 개수만큼 xml 생성
        slide_prompt = f"""
        {content}
        위 내용에서 {project}에 대한 내용을 추출합니다.
        xml_template/ppt/slides의 xml 파일 중 내용과 적절한 xml을 고릅니다.
        이후 선택한 xml에 내용을 적절한 위치에 삽입합니다.
        만약 내용을 모두 채우지 못했다면 채운 요소의 크기를 키우고 나머지를 삭제하여 화면을 최대한 채웁니다. 이때 다른 요소들은 절대 건들지 않고 침범하지 않습니다.
        xml_result/ppt/slides/slide{i}.xml에 저장해주세요
        """
        result = await Runner.run(starting_agent=agent, input=slide_prompt, max_turns=30)
        print(f"Slide{i}의 xml:", result.final_output)
        save_dir = os.path.join(CUR_PATH, 'xml_template')

        new_ppt_name = os.path.join(CUR_PATH, 'new.pptx')
        subprocess.call(f'opc repackage {save_dir} {new_ppt_name}')

        print(new_ppt_name, "에 저장됨")
        # xml_root = ET.fromstring(result.final_output)
        # tree = ET.ElementTree(xml_root)
        # # xml 파일로 저장
        # xml_path = RESULT_XML_PROJECT_PATH.format(slide_idx=i)
        # with open(xml_path, "wb") as xml_file:
        #     tree.write(xml_file, encoding="utf-8", xml_declaration=True)



async def init_servers():
    # Initialize the MCP server with the specified parameters
    # print("Initializing PPT servers...")

    print("Initializing File System servers...")

    server = MCPServerStdio(
        name="Filesystem Server, via npx",
        params={
            "command": "npx",
            "args": ["-y", "@modelcontextprotocol/server-filesystem", "C:/Users/SSAFY/Desktop/repo/github_mcp/ppt/xml"],
        }
    )

    return server

content = """
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

"""


async def main():
    print("PPT servers initialized.")
    async with MCPServerStdio(
        name="Filesystem Server, via npx",
        params={
            "command": "npx",
            "args": ["-y", "@modelcontextprotocol/server-filesystem", "C:/Users/kwon/Desktop/repo/github_mcp/ppt/xml"],
        }
        # prarams={
        #     "command": "docker",
        #     "args": [
        #         "run",
        #         "-i",
        #         "--rm",
        #         "--mount", "type=bind,src=C:/Users/SSAFY/Desktop/repo/github_mcp/ppt/xml,dst=/projects/Desktop",
        #         "--mount", "type=bind,src=C:/Users/SSAFY/Desktop/repo/github_mcp/ppt/xml,dst=/projects/other/allowed/dir,ro",
        #         "mcp/filesystem",
        #         "/projects"
        #         ]
        # }
    ) as s:
        mcp_servers = [s]
        with trace(workflow_name="MCP PPT XML Example"):
            print('t')
            await run(mcp_servers=mcp_servers, content=content)

if __name__ == "__main__":
    # if not shutil.which("uv"):
    #     raise RuntimeError("uv is not installed. Please install it with `pip install uv`.")

    asyncio.run(main())
