---
layout: post
title: "Microsoft Graph Email Sender with OAuth 2.0 in C++"
date: 2024-02-28 10:00:00 +0530
categories: [C++, Windows, Microsoft Graph, OAuth]
tags: [cpp, winhttp, microsoft-graph, oauth2, email, attachments]
author: Arulmurugan K
description: "Complete guide to sending emails with attachments via Microsoft Graph API using OAuth 2.0 in C++ (Visual Studio 2008). Includes smart attachment handling for files up to 150MB."
---

<div align="center">
  <img src="https://img.shields.io/badge/Version-1.0.2-blue.svg" alt="Version">
  <img src="https://img.shields.io/badge/Platform-Windows-lightgrey.svg" alt="Platform">
  <img src="https://img.shields.io/badge/VS-2008-purple.svg" alt="Visual Studio">
  <img src="https://img.shields.io/badge/License-MIT-green.svg" alt="License">
</div>

<br>

<div align="center">
  <h1>📧 Microsoft Graph Email Sender</h1>
  <p><em>A robust C++ implementation for sending emails with attachments via Microsoft Graph API</em></p>
  <p><strong>By Arulmurugan K</strong></p>
  <p><em>Copyright © 2024 Arulmurugan K. All Rights Reserved.</em></p>
</div>

---

## 📋 Table of Contents
- [Introduction](#introduction)
- [Features](#features)
- [Prerequisites](#prerequisites)
- [Architecture](#architecture)
- [Installation](#installation)
- [Azure AD Setup](#azure-ad-setup)
- [Complete Source Code](#complete-source-code)
- [Usage Examples](#usage-examples)
- [How It Works](#how-it-works)
- [Troubleshooting](#troubleshooting)
- [License](#license)
- [About the Author](#about-the-author)

---

## 🎯 Introduction

**Microsoft Graph Email Sender** is a powerful C++ application I developed to enable seamless email sending with attachments through Microsoft Graph API. The application implements OAuth 2.0 authentication and intelligently handles files of all sizes - from tiny text files to large attachments up to 150MB.

### Why I Built This
While working on enterprise automation systems, I needed a reliable way to send emails with attachments from C++ applications. Existing solutions were either too heavy (requiring .NET) or didn't support modern OAuth 2.0 authentication. This led me to create a lightweight, dependency-free solution using native Win32 APIs.

### Perfect For
- ✅ Enterprise automation systems
- ✅ Automated reporting tools
- ✅ Backup notification systems
- ✅ Batch email processing
- ✅ Integration with existing C++ applications
- ✅ Legacy system modernization

---

## ✨ Key Features

| Feature | Description |
|---------|-------------|
| 🔐 **OAuth 2.0 Authentication** | Secure client credentials flow with Azure AD |
| 🔄 **Automatic Token Management** | Token caching and auto-refresh 5 minutes before expiry |
| 📎 **Smart Attachment Handling** | Auto-detects file size and chooses optimal upload method |
| 📦 **Multiple File Support** | Send unlimited attachments per email |
| 🔧 **Visual Studio 2008 Compatible** | Pure Win32 API - no MFC/ATL dependencies |
| ⚡ **High Performance** | Native code with efficient memory management |
| 📝 **MIME Type Detection** | Automatic content-type identification |
| 🛡️ **Error Recovery** | Comprehensive error handling with retry logic |

---

## 📋 Prerequisites

### Development Environment
| Requirement | Specification |
|-------------|--------------|
| **IDE** | Visual Studio 2008 or later |
| **Platform** | Windows XP/Vista/7/8/10/11 |
| **Language** | C++ (Unicode) |
| **SDK** | Windows SDK (includes WinHTTP) |
| **Libraries** | WinHTTP.lib, Crypt32.lib |

### Azure Requirements
- ✅ Active Azure Subscription
- ✅ Azure AD Tenant
- ✅ Registered Application in Azure AD
- ✅ Mail.Send API Permission
- ✅ Client Secret

---

## 🏗 Architecture

### File Size Handling Logic

File Attachment
│
▼
┌─────────────────────────┐
│ Check File Size │
└─────────────────────────┘
│
┌───────────┴───────────┐
▼ ▼
┌────────────────┐ ┌────────────────┐
│ Size ≤ 3MB │ │ Size > 3MB │
└────────────────┘ └────────────────┘
│ │
▼ ▼
┌────────────────┐ ┌────────────────┐
│ Direct Upload │ │ Create Upload │
│ (Base64 JSON) │─────▶│ Session │
└────────────────┘ └────────────────┘
│
▼
┌────────────────┐
│ Upload in │
│ 5MB Chunks │
└────────────────┘

## 🔧 Installation

### Step 1: Create Visual Studio Project

1. Open **Visual Studio 2008**
2. `File → New → Project`
3. Select **Win32 Console Application**
4. Name: `GraphEmailSender`
5. Click **OK**
6. In Application Wizard:
   - Application Type: **Console application**
   - Additional options: **Empty project**
   - Character Set: **Use Unicode Character Set**

### Step 2: Add Source File

1. Right-click **Source Files** folder
2. `Add → New Item`
3. Select **C++ File (.cpp)**
4. Name: `GraphEmailSender.cpp`
5. Click **Add**

### Step 3: Configure Project

1. `Project → Properties`
2. Configuration: **All Configurations**
3. `Linker → Input → Additional Dependencies`: winhttp.lib, crypt32.lib
4. Click **OK**

### Step 4: Copy Source Code

Copy the complete source code from the [Complete Source Code](#complete-source-code) section below.

### Step 5: Build

`Build → Build Solution` (or press **F7**)

---

## 🔑 Azure AD Setup

### Step-by-Step Configuration

#### 1. Register Application
1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory → App registrations**
3. Click **New registration**
4. Name: `Email Sender Application`
5. Supported account types: **"Accounts in this organizational directory only"**
6. Click **Register**

#### 2. Add API Permissions
1. Go to **API permissions**
2. Click **Add a permission**
3. Select **Microsoft Graph**
4. Choose **Application permissions**
5. Search for and select **Mail.Send**
6. Click **Add permissions**
7. Click **Grant admin consent** (requires admin privileges)

#### 3. Create Client Secret
1. Go to **Certificates & secrets**
2. Click **New client secret**
3. Description: `Email Sender Secret`
4. Expiry: **24 months** (recommended)
5. Click **Add**
6. ⚠️ **COPY THE SECRET VALUE NOW** - you won't be able to view it later!

#### 4. Gather Required Values
Client ID: 12345678-1234-1234-1234-123456789012
Tenant ID: 12345678-1234-1234-1234-123456789012
Client Secret: ************************************
User ID: sender@yourcompany.com

---

## 📝 Complete Source Code

<details>
<summary>📁 Click to expand and view the complete source code (650+ lines)</summary>

```cpp
/*****************************************************************************
* MICROSOFT GRAPH EMAIL SENDER
* 
* Copyright (c) 2024 Arulmurugan K
* All Rights Reserved.
*
* This software is the proprietary information of Arulmurugan K.
* Use is subject to license terms.
*
* Author: Arulmurugan K
* Version: 1.0.2
* Date: February 2024
* Platform: Windows (Visual Studio 2008)
* 
* Description: Send emails with attachments via Microsoft Graph API
*              using OAuth 2.0 authentication.
*****************************************************************************/

// PASTE YOUR COMPLETE SOURCE CODE HERE
// (The 650+ lines you shared earlier in our conversation)

// Make sure your copyright header is included:
/*****************************************************************************
* MICROSOFT GRAPH EMAIL SENDER
* 
* Copyright (c) 2024 Arulmurugan K
* All Rights Reserved.
*
* Author: Arulmurugan K
*****************************************************************************/

// [Your complete source code goes here]

🚀 Usage Examples
#include "GraphEmailSender.h"

int main()
{
    // Configuration
    std::wstring clientId = L"12345678-1234-1234-1234-123456789012";
    std::wstring clientSecret = L"your-client-secret";
    std::wstring tenantId = L"12345678-1234-1234-1234-123456789012";
    std::wstring userId = L"reports@company.com";

    // Email details
    std::wstring to = L"manager@company.com";
    std::wstring subject = L"Weekly Sales Report";
    std::wstring body = L"Please find attached the weekly sales reports.";

    // Attachments
    std::vector<std::wstring> attachments;
    attachments.push_back(L"C:\\reports\\sales_q1_2024.xlsx");
    attachments.push_back(L"C:\\reports\\summary.pdf");

    // Send email
    EmailSender sender(clientId, clientSecret, tenantId, userId);
    if (sender.SendEmail(to, subject, body, attachments))
    {
        wprintf(L"✓ Weekly report sent successfully!\n");
    }
    else
    {
        wprintf(L"✗ Failed to send report.\n");
    }

    return 0;
}

Example 2: Sending Large Backup Files
// This handles files > 3MB automatically with chunked upload
std::vector<std::wstring> attachments;
attachments.push_back(L"D:\\backups\\database_backup_2024.bak");  // 500MB
attachments.push_back(L"D:\\backups\\logs_archive.zip");          // 200MB

if (sender.SendEmail(L"admin@company.com", 
                     L"Daily Backup Files", 
                     L"Database and log backups attached",
                     attachments))
{
    wprintf(L"✓ Large backup files sent successfully!\n");
}

🔍 How It Works

1. OAuth 2.0 Token Flow
Application → Azure AD Token Endpoint → Access Token → Microsoft Graph API
2. Email Creation Flow

    Create Draft Message - Creates a new email draft

    Upload Attachments - Adds files based on size

    Send Message - Sends the completed email

3. Smart Attachment Logic
if (fileSize <= MAX_SMALL_FILE_SIZE) {
    // Direct Base64 upload for files ≤ 3MB
    AddSmallAttachment(messageId, filePath);
} else {
    // Chunked upload for files > 3MB
    AddLargeAttachment(messageId, filePath);
}

4. Chunked Upload Process

    Create Upload Session - Get upload URL from Graph API

    Upload Chunks - Send 5MB chunks with Content-Range headers

    Complete - Last chunk completes the upload

❗ Troubleshooting
Common Issues and Solutions
Error	Cause	Solution
401 - Unauthorized :	Token expired or invalid	Check client secret, ensure token refresh logic is working
403 - Forbidden	Insufficient permissions: 	Verify Mail.Send permission is granted and admin consented
413 - Payload Too Large	File too large for direct upload: 	Use chunked upload (automatic in code)
File not found - 	Invalid file path: Check file paths and permissions

// Enable HTTP tracing
DWORD statusCode = 0;
DWORD statusCodeSize = sizeof(statusCode);
WinHttpQueryHeaders(hRequest, 
    WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
    NULL, &statusCode, &statusCodeSize, NULL);
wprintf(L"HTTP Status: %d\n", statusCode);


📄 License
MIT License

Copyright (c) 2024 Arulmurugan K

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

👨‍💻 About the Author

<div align="center"> <h3>Arulmurugan K</h3> <p><em>Software Developer | C++ Specialist | Cloud Enthusiast</em></p> <p> I'm a passionate software developer with expertise in C++ programming, Windows application development, and cloud integration. I specialize in creating efficient and secure applications that leverage modern cloud services. </p> <h4>Expertise</h4> <ul style="list-style: none; padding: 0;"> <li>🔹 <strong>Programming Languages:</strong> C++, C#, Python</li> <li>🔹 <strong>Technologies:</strong> Microsoft Graph API, OAuth 2.0, REST APIs</li> <li>🔹 <strong>Platforms:</strong> Windows, Azure Cloud</li> <li>🔹 <strong>Tools:</strong> Visual Studio, Git, WinHTTP</li> </ul> <h4>Connect with Me</h4> <p> 📧 Email: arulmurugan@example.com<br> 📝 GitHub: <a href="https://github.com/arulmurugank">https://github.com/arulmurugank</a> </p> </div>

📊 Project Statistics

Metric	       |     Value
-----------------------------------------
Lines of Code	 |     ~650
Classes        |	    2
Functions	     |     12
Files	         |      1
Dependencies	 |     WinHTTP, Crypt32
Compatibility	 |   Windows XP to Windows 11
-----------------------------------------------

