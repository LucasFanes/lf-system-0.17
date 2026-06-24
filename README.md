# LF System

A Python-based administrative system designed for account management, purchase tracking, spreadsheet generation, and Google Sheets integration.

## Overview

LF System is a personal project developed to improve software engineering, database management, file handling, and automation skills.

The application provides a set of tools for managing customer accounts, tracking purchases, generating reports, and performing backup operations through a terminal-based interface.

## Features

### Account Management

* Create customer accounts
* Edit account information
* View account details
* Monitor registered customers

### Purchase Management

* Register purchases
* Track purchase history
* Delete entries using unique NC codes
* Monitor global transaction history

### Spreadsheet Operations

* Generate Excel reports automatically
* Upload spreadsheets to Google Sheets
* Download Google Sheets in multiple formats:

  * XLSX
  * CSV
  * PDF
  * ODS
  * TSV
  * ZIP

### File and Backup Utilities

* ZIP backup generation
* File replication
* Folder replication
* Safe deletion (Recycle Bin)
* Permanent deletion with confirmation

### Database

* SQLite-based persistent storage
* Automatic schema creation
* Legacy database migration support
* Account and purchase relationship management

## Technologies Used

* Python 3
* SQLite
* OpenPyXL
* EZSheets
* PyMsgBox
* Send2Trash

## Project Structure

```text
LF-System/
│
├── Main.py
├── Billing.py
├── Function.py
├── Logs.py
├── VERIFICAR_BANCO.py
│
├── Data/
├── Backups/
├── Logs/
└── Spreadsheets/
```

## Database Schema

### Accounts Table

| Field         | Description           |
| ------------- | --------------------- |
| name          | Account name          |
| code          | Internal account code |
| created_at    | Creation timestamp    |
| billing_items | Customer metadata     |

### Purchases Table

| Field        | Description         |
| ------------ | ------------------- |
| nc           | Purchase identifier |
| account_name | Related account     |
| item         | Purchased item      |
| price        | Purchase value      |
| date         | Purchase timestamp  |

## Installation

Clone the repository:

```bash
git clone https://github.com/your-username/LF-System.git
cd LF-System
```

Install dependencies:

```bash
pip install pyperclip
pip install pymsgbox
pip install send2trash
pip install openpyxl
pip install ezsheets
```

Or:

```bash
pip install -r requirements.txt
```

## Running the Application

```bash
python Main.py
```

## Learning Objectives

This project was built to practice and improve knowledge in:

* Object-Oriented Programming (OOP)
* Database Design
* SQLite Integration
* Spreadsheet Automation
* File Management
* Software Architecture
* Python Development
* Logging and Monitoring Systems

## Future Improvements

* Graphical User Interface (GUI)
* User Authentication System
* Dashboard and Analytics
* Cloud Database Integration
* REST API Support
* Multi-user Environment
* Web Version

## Author

Lucas Fanes

Computer Science and Computer Engineering Student at FMU.

Currently seeking internship and junior software development opportunities to apply and expand skills in programming, automation, and software engineering.

## License

This project is intended for educational and portfolio purposes.
