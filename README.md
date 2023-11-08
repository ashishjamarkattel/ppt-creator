# ppt-creator
[![Python 3.8](https://img.shields.io/badge/python-3.6-blue.svg)](https://www.python.org/downloads/release/python-360/)
![License](https://img.shields.io/github/license/AI4Finance-Foundation/fingpt.svg?color=brightgreen)

Enough of boring way to manually create powerpoint, this ppt-creator, directly creates the powerpoints so you can easily make changes to them or finish it within powerpoint.
It also have placeholders for images!, to make your powerpoint attractive.

## 🔧 Quick start
```bash
pip install streamlit openai python-pptx google-serp-api

streamlit run app.py
```
# How it works:
- The user sends a name of ppt to create along with username
- The GPT 3.5 model creates the content for the ppt.
- Google-serp api then download the image using specific header.
- The Python-pptx library converts the generated content into a PowerPoint presentation..
- This tool is perfect for anyone who wants to quickly create professional-looking PowerPoint presentations without spending hours on design and content creation.

## 📝 What does it create ?
![alt text](https://github.com/ashishjamarkattel/ppt-creator/blob/master/assets/cow-ppt.PNG)
