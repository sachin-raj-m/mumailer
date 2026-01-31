# How to Run & Deploy MuMailer Web

## 1. Run Locally (on your computer)

Since you already have Streamlit installed, just run:

```bash
streamlit run app.py
```

This will open the app in your browser (usually at `http://localhost:8501`).

## 2. Deploy to the Web (for everyone to use)

The easiest way to host this for free is using **Streamlit Community Cloud**.

### Prerequisites
1.  Push your code to a GitHub repository (it seems you are already in a git repo).
2.  Make sure `requirements.txt` is in the root folder (I have already updated it).

### Steps
1.  Go to [share.streamlit.io](https://share.streamlit.io/).
2.  Sign in with GitHub.
3.  Click **"New App"**.
4.  Select your GitHub Repository, Branch (usually `main`), and File Path (`app.py`).
5.  Click **"Deploy"**.

That's it! You will get a unique URL (e.g., `https://mumailer.streamlit.app`) that you can share with your team. They can access it from any device without installing Python.
