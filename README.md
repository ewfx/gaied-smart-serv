# 🚀 Project Name

## 📌 Table of Contents
- [Introduction](#introduction)
- [Demo](#demo)
- [Inspiration](#inspiration)
- [What It Does](#what-it-does)
- [How We Built It](#how-we-built-it)
- [How to Run](#how-to-run)
- [Tech Stack](#tech-stack)
- [Team](#team)

---

## 🎯 Introduction
Generative AI-powered solution to automate the processing of email data and attachments to identify critical information, such as request types and Sub-Request Type, enhancing efficiency, accuracy, and turnaround time. 

The solution must minimize manual interventions by leveraging freely available tools to classify emails, extract relevant contextual information, and create service requests

## 🎥 Demo
Video Recording => https://www.loom.com/share/8acce5d9e5194bb0bd1fe4ad665b1329


## 💡 Inspiration
The inspiration behind this project stems from the challenges organizations face in managing large volumes of email-based service requests. Manually classifying emails, extracting relevant information, and routing them to the right teams can be time-consuming, error-prone, and inefficient.

This initiative aims to address these challenges by leveraging Generative AI to create an optimized system that automates these processes.

## ⚙️ What It Does

**Email Classification:** 
Utilizes Generative AI to categorize emails into predefined request types and subtypes, accurately interpreting sender intent.

**Contextual Data Extraction:** 
Extracts configurable fields from both email bodies and attachments like PDFs and images by leveraging advanced AI-driven text interpretation techniques (Gemini Gen AI model)

**Automated Service Request Creation:** 
Transfers extracted information into service requests, streamlining the process for easy tracking and management.

**Efficiency and Accuracy:** 
Reduces manual intervention, enhances data processing accuracy, and accelerates turnaround times.

## 🛠️ How We Built It

**Generative AI Models:** 
Advanced AI model (Google Gemini AI) for email classification, data extraction, and contextual text analysis. 

**Programming Language:** 
Python, for implementing logic and integrating tools.

**Text Processing Tools:** 
Libraries like pytesseract, pypdf2, and BytesIO for extracting and processing text from email bodies and attachments.

**Email Handling:** 
Gmail IMAP Server for accessing and fetching unread emails.

**Data Management:** 
Service requests are created in Google Sheets to effectively record and organize extracted data.


## 🏃 How to Run
Code is provided under https://github.com/ewfx/gaied-smart-serv/blob/main/code/src/main.py

User need to Run this code by installing required python libraries and also need to provide Gemini API Key for integrating Gemini GenAI.

This will create a service request record in google sheets (Ex: https://docs.google.com/spreadsheets/d/1yLsOlV0Z3_JXjsRcAlSisIycs1MFJ-xzI8fUa2Xd0eI/edit?gid=638029294#gid=638029294) and also send acknowledgement email to the sender.

## 🏗️ Tech Stack
- 🔹 Backend: Flask API
- 🔹 Other: GEMINI API

## 👥 Team
Tribhuvan Kumar
Rahul
Anand
Abinash


