# 📞 CRM Call Tracking System

A lightweight **CRM Call Tracking System** built using **Google Apps Script**, with **Google Sheets as the backend database**.

This project is designed to manage customer/party calling activities, maintain call history, track follow-ups, and organize calling data in a structured and centralized system.

The main objective of this project is to demonstrate how **Google Apps Script can be used to convert a manual business process into a simple, automated, and data-driven application**.

---

## 🎯 Project Objective

In many organizations, calling activities are managed manually using Excel or Google Sheets. While these tools are useful for data management, repetitive manual processes can create several challenges.

### Common Business Problems

- Difficulty tracking previous calls
- Missed follow-ups
- Repetitive data entry
- Duplicate records
- No centralized calling history
- Difficulty monitoring employee/agent performance
- Manual report preparation
- Limited visibility into pending follow-ups
- Data inconsistency

### Proposed Solution

This project provides a simple CRM-style calling system that centralizes calling information and allows users to manage customer interactions, call outcomes, remarks, and follow-up activities in one place.

---

# 🚀 Key Features

## 👤 Customer / Party Management

- Store customer/party information
- Search customer records
- View customer details
- Maintain customer calling history
- Track customer-related information

## 📞 Call Tracking

- Record call details
- Capture call date and time
- Select call status/outcome
- Add call remarks
- Maintain complete calling history
- Track the latest customer interaction

## 🔄 Follow-up Management

- Add follow-up dates
- Track pending follow-ups
- Identify upcoming follow-ups
- Maintain follow-up history
- Track completed and pending activities

## 🔍 Search & Data Retrieval

- Search customer/party records
- Retrieve existing information
- View previous call activities
- Access relevant calling history

## ⚙️ Automation

Google Apps Script is used to automate application logic and interaction with Google Sheets.

The system can perform tasks such as:

- Adding records
- Updating records
- Searching records
- Reading data
- Writing data
- Maintaining call history
- Processing user inputs
- Automating repetitive operations

---

# 🏗️ Technology Stack

| Technology | Purpose |
|---|---|
| Google Apps Script | Backend logic and automation |
| JavaScript | Application and business logic |
| HTML | User interface |
| CSS | UI design and styling |
| Google Sheets | Backend database / data storage |
| Google Workspace | Application environment |

---

# 🧩 System Architecture

```text
                ┌─────────────────────┐
                │      User / Agent   │
                └──────────┬──────────┘
                           │
                           ▼
                ┌─────────────────────┐
                │   CRM Web Interface │
                │      HTML / CSS     │
                └──────────┬──────────┘
                           │
                           ▼
                ┌─────────────────────┐
                │  Google Apps Script │
                │   Business Logic    │
                └──────────┬──────────┘
                           │
                           ▼
                ┌─────────────────────┐
                │    Google Sheets    │
                │   Backend Database  │
                └─────────────────────┘
