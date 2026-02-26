# 🧾 BillExpenditureCalculator – Smart Bill Scanner & Expense Tracker

A desktop-based intelligent bill scanning system that extracts total amounts from receipts using OCR and automatically generates structured Excel reports. Built with Python, this application integrates computer vision, text recognition, real-time camera scanning, voice feedback, and interactive GUI components.

## Features
## Image & Real-Time Bill Scanning

Upload bill images manually

Capture bills using real-time camera scanning

Video-based animated intro interface

## OCR-Based Text Extraction

Uses Tesseract OCR to extract text from receipts

Regex-based parsing to identify total bill amount

User validation before saving extracted amount

## Automated Excel Reporting

Stores bill number and total amount in Excel

Auto-formats columns (alignment, borders, fonts)

Generates total expenditure report

Automatically opens Excel report file

## Interactive User Experience

Button click sound effects

Background music during intro

Text-to-speech validation feedback

Visual success/failure indicators

## Tech Stack

Python

Tkinter – GUI Development

OpenCV – Camera & Video Processing

PyTesseract – OCR Text Extraction

OpenPyXL – Excel Automation

Pygame – Audio Integration

pyttsx3 – Text-to-Speech

Pillow (PIL) – Image Handling

## Project Workflow

Launch application (fullscreen mode)

Intro video + background music plays

Choose:

Scan image from system

Real-time camera scan

OCR extracts text from bill

System detects total amount

User confirms extracted value

Data stored in Excel

Generate expenditure report

## Key Implementation Highlights

Multi-threaded video playback for smooth UI

Regex-based intelligent total detection

Real-time webcam capture with OpenCV

Styled Excel automation with formatting

Event-driven GUI design

Audio + voice-assisted feedback system

## File Requirements

Before running the project, ensure:

Tesseract OCR installed and configured

Required media files (click sound, intro music, video, icons)

Excel installed (for report auto-opening)

## How to Run
pip install pytesseract opencv-python pillow openpyxl pyttsx3 pygame

Then run:

python main.py
## Future Improvements

Cloud-based bill storage

Expense categorization (Food, Travel, Shopping)

Monthly analytics dashboard

Database integration

Cross-platform packaging (.exe)

## Role

Team Lead

Designed system architecture

Implemented OCR + Regex logic

Integrated Excel automation

Built full GUI workflow
