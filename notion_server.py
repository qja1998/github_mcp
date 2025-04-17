import asyncio
import shutil

import os
import dotenv
dotenv.load_dotenv()

from collections import deque

GITHUB_PERSONAL_ACCESS_TOKEN = os.getenv("GITHUB_PERSONAL_ACCESS_TOKEN")
OPENAI_MODEL = os.getenv("OPENAI_MODEL")

from agents import Agent, Runner, trace
from agents.mcp import MCPServer, MCPServerStdio

async def run(mcp_server: MCPServer):
    agent = Agent(
        model=OPENAI_MODEL,
        name="Assistant",
        instructions=f"Answer questions about the notion documents",
        mcp_servers=[mcp_server],
    )

    PROMPT = f"""
    You are a potfolio assistant. You are given the user's notion page.
    철저하게 사용자 관점에서 답변하라.
    You should answer questions about the notion documents.
    NEVER answer questions about the directories that are not in the notion documents.
    답변은 절대 지어내지 말고, 항상 사실에 기반해야 한다.
    """

    chat_history = deque([], maxlen=5)

    while True:
        # Ask the user for the git command
        command = input("Please enter a git command (or 'exit' to quit): ")
        if command.lower() == "exit":
            break
        
        END_PROMPT = f"""
        history: {chat_history}
        user: {command}"""
        prompt = PROMPT.format(chat_history=chat_history) + END_PROMPT
        print(prompt)

        # Run the command and print the result
        print("\n" + "-" * 40)
        print(f"Running: {command}")
        result = await Runner.run(starting_agent=agent, input=prompt)
        print(result.final_output)
        current_chat = [
            {"role": "user", "content": command},
            {"role": "assistant", "content": result.final_output},
        ]
        chat_history.append(current_chat)
        print("\n" + "-" * 40)


async def main():
    # Ask the user for the directory path

    async with MCPServerStdio(
        cache_tools_list=True,  # Cache the tools list, for demonstration
        params={
            "command": "python",
            "args": ["-m", "notion_mcp"],
        }
    ) as server:
        print("setted notion server")
        with trace(workflow_name="MCP Notion Example"):
            await run(server)


if __name__ == "__main__":
    if not shutil.which("uvx"):
        raise RuntimeError("uvx is not installed. Please install it with `pip install uvx`.")

    asyncio.run(main())