from dotenv import load_dotenv
from langchain import hub
from langchain.agents import AgentExecutor, create_react_agent
from langchain_community.tools import DuckDuckGoSearchRun
from langchain_groq import ChatGroq

load_dotenv()

search_tool = DuckDuckGoSearchRun()
tools = [search_tool]

llm = ChatGroq(
    model="qwen/qwen3.8-27b",
    temperature=0,
)

prompt = hub.pull("hwchase17/react")
agent = create_react_agent(
    llm=llm,
    tools=tools,
    prompt=prompt,
)

agent_executor = AgentExecutor(
    agent=agent,
    tools=tools,
    verbose=True,
    max_iterations=5,
)


if __name__ == "__main__":
    response = agent_executor.invoke(
        {"input": "What is the release date of the movie Dhadak 2?"}
    )
    print(response["output"])
