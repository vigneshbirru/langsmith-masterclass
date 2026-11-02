
import os
from dotenv import load_dotenv

from langchain_community.document_loaders import PyPDFLoader
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_huggingface import HuggingFaceEmbeddings
from langchain_groq import ChatGroq
from langchain_community.vectorstores import FAISS
from langchain_core.prompts import ChatPromptTemplate
from langchain_core.runnables import (
    RunnableParallel,
    RunnablePassthrough,
    RunnableLambda,
)
from langchain_core.output_parsers import StrOutputParser

os.environ['LANGCHAIN_PROJECT'] = 'RAG CHATBOT'
# Load environment variables
load_dotenv()

if not os.getenv("GROQ_API_KEY"):
    raise ValueError("Please set GROQ_API_KEY in your .env file")

PDF_PATH = "islr.pdf"

# 1. Load PDF
loader = PyPDFLoader(PDF_PATH)
docs = loader.load()

# 2. Split documents into chunks
splitter = RecursiveCharacterTextSplitter(
    chunk_size=1000,
    chunk_overlap=150,
)
splits = splitter.split_documents(docs)

# 3. Hugging Face embeddings (runs locally)
emb = HuggingFaceEmbeddings(
    model_name="sentence-transformers/all-MiniLM-L6-v2",
    model_kwargs={"device": "cpu"},
    encode_kwargs={"normalize_embeddings": True},
)

# 4. Create FAISS vector index
vs = FAISS.from_documents(splits, emb)

retriever = vs.as_retriever(
    search_type="similarity",
    search_kwargs={"k": 4},
)

# 5. Prompt template
prompt = ChatPromptTemplate.from_messages([
    (
        "system",
        "Answer only using the provided context. "
        "If the answer is not present in the context, "
        "say you don't know. Do not invent facts."
    ),
    (
        "human",
        "Question: {question}\n\nContext:\n{context}"
    ),
])

# 6. Groq LLM with Qwen     3.8 27B
llm = ChatGroq(
    model="qwen/qwen3.8-27b",
    temperature=0,
    reasoning_effort="none",
)

# Format retrieved documents
def format_docs(docs):
    return "\n\n".join(doc.page_content for doc in docs)

# 7. Build RAG chain
parallel = RunnableParallel({
    "context": retriever | RunnableLambda(format_docs),
    "question": RunnablePassthrough(),
})

chain = parallel | prompt | llm | StrOutputParser()

# 8. Ask questions interactively
print("PDF RAG ready! Ask questions about your PDF.")
print("Type 'exit' to quit.")

while True:
    q = input("\nQ: ").strip()

    if q.lower() in {"exit", "quit"}:
        break

    if not q:
        continue

    try:
        ans = chain.invoke(q)
        print("\nA:", ans)
    except Exception as e:
        print(f"\nError: {e}")
