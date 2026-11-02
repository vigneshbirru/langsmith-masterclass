# Install required packages:
# pip install -U langchain langchain-openai langchain-community langchain-huggingface langchain-groq faiss-cpu pypdf python-dotenv langsmith sentence-transformers

import os
from dotenv import load_dotenv
from langsmith import traceable

from langchain_community.document_loaders import PyPDFLoader
from langchain.text_splitter import RecursiveCharacterTextSplitter
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

# ---------- Environment setup ----------
load_dotenv()

os.environ["LANGCHAIN_PROJECT"] = "RAG CHATBOT"

# Set these in your .env file:
# LANGCHAIN_TRACING_V2=true
# LANGCHAIN_API_KEY=your_langsmith_api_key
# GROQ_API_KEY=your_groq_api_key

PDF_PATH = "islr.pdf"


# ---------- PDF loading ----------
@traceable(
    name="load_pdf",
    tags=["pdf", "loader"],
    metadata={"loader": "PyPDFLoader"},
)
def load_pdf(path: str):
    if not os.path.exists(path):
        raise FileNotFoundError(f"PDF file not found: {path}")

    loader = PyPDFLoader(path)
    return loader.load()


# ---------- Document splitting ----------
@traceable(name="split_documents")
def split_documents(docs, chunk_size=1000, chunk_overlap=150):
    if not docs:
        raise ValueError("No documents were loaded from the PDF.")

    splitter = RecursiveCharacterTextSplitter(
        chunk_size=chunk_size,
        chunk_overlap=chunk_overlap,
    )

    return splitter.split_documents(docs)


# ---------- Vector store creation ----------
@traceable(
    name="build_vectorstore",
    tags=["embedding", "vectorstore"],
    metadata={"embedding_model": "all-MiniLM-L6-v2"},
)
def build_vectorstore(splits):
    if not splits:
        raise ValueError("No document chunks were created.")

    embeddings = HuggingFaceEmbeddings(
        model_name="sentence-transformers/all-MiniLM-L6-v2",
        model_kwargs={"device": "cpu"},
        encode_kwargs={"normalize_embeddings": True},
    )

    vectorstore = FAISS.from_documents(splits, embeddings)
    return vectorstore


# ---------- Setup pipeline ----------
@traceable(name="setup_pipeline")
def setup_pipeline(pdf_path: str):
    docs = load_pdf(pdf_path)
    splits = split_documents(docs)
    vectorstore = build_vectorstore(splits)
    return vectorstore


# ---------- LLM ----------
llm = ChatGroq(
    model="openai/gpt-oss-120b",
    temperature=0,
)


# ---------- Prompt ----------
prompt = ChatPromptTemplate.from_messages(
    [
        (
            "system",
            "Answer only using the provided context. "
            "If the answer is not present in the context, "
            "say you don't know. Do not make up information.",
        ),
        (
            "human",
            "Question: {question}\n\nContext:\n{context}",
        ),
    ]
)


# ---------- Format retrieved documents ----------
def format_docs(docs):
    return "\n\n".join(doc.page_content for doc in docs)


# ---------- Build RAG chain ----------
vectorstore = setup_pipeline(PDF_PATH)

retriever = vectorstore.as_retriever(
    search_type="similarity",
    search_kwargs={"k": 4},
)

parallel = RunnableParallel(
    {
        "context": retriever | RunnableLambda(format_docs),
        "question": RunnablePassthrough(),
    }
)

chain = parallel | prompt | llm | StrOutputParser()


# ---------- Ask questions ----------
def main():
    print("PDF RAG chatbot ready!")
    print("Ask a question about your PDF, or type 'exit' to quit.")

    while True:
        try:
            question = input("\nQ: ").strip()

            if question.lower() in {"exit", "quit"}:
                print("Goodbye!")
                break

            if not question:
                continue

            config = {
                "run_name": "pdf_rag_query",
                "tags": ["rag", "pdf"],
                "metadata": {"source": PDF_PATH},
            }

            answer = chain.invoke(question, config=config)
            print("\nA:", answer)

        except KeyboardInterrupt:
            print("\nGoodbye!")
            break

        except Exception as e:
            print(f"\nError: {e}")


if __name__ == "__main__":
    main()