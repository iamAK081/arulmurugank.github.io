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

#include <windows.h>
#include <winhttp.h>
#include <stdio.h>
#include <string>
#include <vector>
#include <algorithm>
#include <ctime>

#pragma comment(lib, "winhttp.lib")
#pragma comment(lib, "crypt32.lib")

// Constants
#define CHUNK_SIZE (5 * 1024 * 1024) // 5MB chunks for large files
#define MAX_SMALL_FILE_SIZE (3 * 1024 * 1024) // 3MB limit for direct attachment
#define TOKEN_EXPIRY_BUFFER 300 // 5 minutes buffer for token expiry

// Helper functions
namespace Utils
{
    // Convert ANSI string to wide string
    std::wstring StringToWString(const std::string& s)
    {
        int len = MultiByteToWideChar(CP_ACP, 0, s.c_str(), (int)s.length(), NULL, 0);
        std::wstring ws(len, 0);
        MultiByteToWideChar(CP_ACP, 0, s.c_str(), (int)s.length(), &ws[0], len);
        return ws;
    }

    // Convert wide string to ANSI string
    std::string WStringToString(const std::wstring& ws)
    {
        int len = WideCharToMultiByte(CP_ACP, 0, ws.c_str(), (int)ws.length(), NULL, 0, NULL, NULL);
        std::string s(len, 0);
        WideCharToMultiByte(CP_ACP, 0, ws.c_str(), (int)ws.length(), &s[0], len, NULL, NULL);
        return s;
    }

    // Read file into byte vector
    std::vector<BYTE> ReadFileA(const std::wstring& filePath)
    {
        HANDLE hFile = CreateFile(filePath.c_str(), GENERIC_READ, FILE_SHARE_READ, NULL,
            OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
        if (hFile == INVALID_HANDLE_VALUE)
            return std::vector<BYTE>();

        DWORD fileSize = GetFileSize(hFile, NULL);
        if (fileSize == INVALID_FILE_SIZE)
        {
            CloseHandle(hFile);
            return std::vector<BYTE>();
        }

        std::vector<BYTE> buffer(fileSize);
        DWORD bytesRead;
        if (ReadFile(hFile, &buffer[0], fileSize, &bytesRead, NULL) == FALSE)
        {
            CloseHandle(hFile);
            return std::vector<BYTE>();
        }

        CloseHandle(hFile);
        return buffer;
    }

    // Base64 encode binary data
    std::wstring Base64Encode(const std::vector<BYTE>& data)
    {
        DWORD base64Len = 0;
        if (CryptBinaryToStringW(&data[0], (DWORD)data.size(),
            CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF,
            NULL, &base64Len) == FALSE)
            return L"";

        std::wstring result(base64Len, 0);
        if (CryptBinaryToStringW(&data[0], (DWORD)data.size(),
            CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF,
            &result[0], &base64Len) == FALSE)
            return L"";

        // Remove null terminator if present
        if (!result.empty() && result[result.size() - 1] == L'\0')
            result.resize(result.size() - 1);

        return result;
    }

    // Get MIME type based on file extension
    std::wstring GetMimeType(const std::wstring& filename)
    {
        size_t dotPos = filename.find_last_of(L".");
        if (dotPos == std::wstring::npos)
            return L"application/octet-stream";

        std::wstring ext = filename.substr(dotPos + 1);
        std::transform(ext.begin(), ext.end(), ext.begin(), towlower);

        if (ext == L"txt") return L"text/plain";
        if (ext == L"pdf") return L"application/pdf";
        if (ext == L"doc") return L"application/msword";
        if (ext == L"docx") return L"application/vnd.openxmlformats-officedocument.wordprocessingml.document";
        if (ext == L"xls") return L"application/vnd.ms-excel";
        if (ext == L"xlsx") return L"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        if (ext == L"jpg" || ext == L"jpeg") return L"image/jpeg";
        if (ext == L"png") return L"image/png";
        if (ext == L"zip") return L"application/zip";

        return L"application/octet-stream";
    }

    // Extract filename from full path
    std::wstring GetFileName(const std::wstring& filePath)
    {
        size_t lastSlash = filePath.find_last_of(L"\\/");
        return (lastSlash == std::wstring::npos) ? filePath : filePath.substr(lastSlash + 1);
    }
}

// Token Manager Class - Handles OAuth 2.0 authentication
class TokenManager
{
private:
    std::wstring m_token;
    time_t m_expiry;

    bool RequestNewToken(const std::wstring& clientId, const std::wstring& clientSecret,
        const std::wstring& tenantId, const std::wstring& scope)
    {
        HINTERNET hSession = WinHttpOpen(L"TokenClient/1.0", WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
            WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0);
        if (hSession == NULL) return false;

        HINTERNET hConnect = WinHttpConnect(hSession, L"login.microsoftonline.com",
            INTERNET_DEFAULT_HTTPS_PORT, 0);
        if (hConnect == NULL)
        {
            WinHttpCloseHandle(hSession);
            return false;
        }

        std::wstring path = L"/" + tenantId + L"/oauth2/v2.0/token";
        HINTERNET hRequest = WinHttpOpenRequest(hConnect, L"POST", path.c_str(), NULL,
            WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES,
            WINHTTP_FLAG_SECURE);
        if (hRequest == NULL)
        {
            WinHttpCloseHandle(hConnect);
            WinHttpCloseHandle(hSession);
            return false;
        }

        // Prepare URL-encoded form data
        std::wstring formData = L"client_id=" + clientId +
            L"&client_secret=" + clientSecret +
            L"&scope=" + scope +
            L"&grant_type=client_credentials";

        // Add headers
        std::wstring headers = L"Content-Type: application/x-www-form-urlencoded\r\n";
        WinHttpAddRequestHeaders(hRequest, headers.c_str(), (DWORD)headers.length(), WINHTTP_ADDREQ_FLAG_ADD);

        // Send request
        if (WinHttpSendRequest(hRequest,
            WINHTTP_NO_ADDITIONAL_HEADERS,
            0,
            (LPVOID)formData.c_str(),
            (DWORD)formData.length() * sizeof(wchar_t),
            (DWORD)formData.length() * sizeof(wchar_t),
            0) == FALSE)
        {
            WinHttpCloseHandle(hRequest);
            WinHttpCloseHandle(hConnect);
            WinHttpCloseHandle(hSession);
            return false;
        }

        if (WinHttpReceiveResponse(hRequest, NULL) == FALSE)
        {
            WinHttpCloseHandle(hRequest);
            WinHttpCloseHandle(hConnect);
            WinHttpCloseHandle(hSession);
            return false;
        }

        DWORD size = 0, downloaded = 0;
        std::string response;
        do
        {
            size = 0;
            if (WinHttpQueryDataAvailable(hRequest, &size) == FALSE) break;

            std::vector<char> buffer(size + 1);
            if (WinHttpReadData(hRequest, &buffer[0], size, &downloaded) == FALSE) break;
            response.append(&buffer[0], size);
        } while (size > 0);

        WinHttpCloseHandle(hRequest);
        WinHttpCloseHandle(hConnect);
        WinHttpCloseHandle(hSession);

        // Parse token response
        size_t tokenPos = response.find("access_token");
        if (tokenPos == std::string::npos) return false;

        size_t start = response.find("\"", tokenPos + 14) + 1;
        size_t end = response.find("\"", start);
        m_token = Utils::StringToWString(response.substr(start, end - start));

        // Parse expiry time
        size_t expiryPos = response.find("expires_in");
        if (expiryPos != std::string::npos)
        {
            start = response.find(":", expiryPos) + 1;
            end = response.find_first_of(",}", start);
            int expiresIn = atoi(response.substr(start, end - start).c_str());
            m_expiry = time(NULL) + expiresIn;
        }
        else
        {
            m_expiry = time(NULL) + 3600; // Default to 1 hour
        }

        return true;
    }

public:
    TokenManager() : m_expiry(0) {}

    bool GetAccessToken(const std::wstring& clientId, const std::wstring& clientSecret,
        const std::wstring& tenantId, const std::wstring& scope)
    {
        if (m_token.empty() || time(NULL) > (m_expiry - TOKEN_EXPIRY_BUFFER))
        {
            return RequestNewToken(clientId, clientSecret, tenantId, scope);
        }
        return true;
    }

    const std::wstring& Token() const { return m_token; }
};

// Main EmailSender Class
class EmailSender
{
private:
    TokenManager m_tokenManager;
    std::wstring m_userId;

    bool CreateDraftMessage(const std::wstring& to, const std::wstring& subject,
        const std::wstring& body, std::wstring& messageId)
    {
        HINTERNET hSession = WinHttpOpen(L"EmailClient/1.0", WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
            WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0);
        if (hSession == NULL) return false;

        HINTERNET hConnect = WinHttpConnect(hSession, L"graph.microsoft.com",
            INTERNET_DEFAULT_HTTPS_PORT, 0);
        if (hConnect == NULL)
        {
            WinHttpCloseHandle(hSession);
            return false;
        }

        std::wstring path = L"/v1.0/users/" + m_userId + L"/messages";
        HINTERNET hRequest = WinHttpOpenRequest(hConnect, L"POST", path.c_str(), NULL,
            WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES,
            WINHTTP_FLAG_SECURE);
        if (hRequest == NULL)
        {
            WinHttpCloseHandle(hConnect);
            WinHttpCloseHandle(hSession);
            return false;
        }

        std::wstring headers = L"Authorization: Bearer " + m_tokenManager.Token() + L"\r\n"
            L"Content-Type: application/json\r\n";
        WinHttpAddRequestHeaders(hRequest, headers.c_str(), (DWORD)headers.length(), WINHTTP_ADDREQ_FLAG_ADD);

        std::wstring json = L"{\"message\":{"
            L"\"subject\":\"" + subject + L"\","
            L"\"body\":{\"contentType\":\"Text\",\"content\":\"" + body + L"\"},"
            L"\"toRecipients\":[{\"emailAddress\":{\"address\":\"" + to + L"\"}}]"
            L"},"
            L"\"saveToSentItems\":\"true\"}";

        if (WinHttpSendRequest(hRequest, NULL, 0, (LPVOID)json.c_str(),
            (DWORD)json.length() * sizeof(wchar_t),
            (DWORD)json.length() * sizeof(wchar_t), 0) == FALSE)
        {
            WinHttpCloseHandle(hRequest);
            WinHttpCloseHandle(hConnect);
            WinHttpCloseHandle(hSession);
            return false;
        }

        if (WinHttpReceiveResponse(hRequest, NULL) == FALSE)
        {
            WinHttpCloseHandle(hRequest);
            WinHttpCloseHandle(hConnect);
            WinHttpCloseHandle(hSession);
            return false;
        }

        DWORD size = 0, downloaded = 0;
        std::string response;
        do
        {
            size = 0;
            if (WinHttpQueryDataAvailable(hRequest, &size) == FALSE) break;

            std::vector<char> buffer(size + 1);
            if (WinHttpReadData(hRequest, &buffer[0], size, &downloaded) == FALSE) break;
            response.append(&buffer[0], size);
        } while (size > 0);

        WinHttpCloseHandle(hRequest);
        WinHttpCloseHandle(hConnect);
        WinHttpCloseHandle(hSession);

        // Parse message ID
        size_t idPos = response.find("id");
        if (idPos == std::string::npos) return false;

        size_t start = response.find("\"", idPos + 4) + 1;
        size_t end = response.find("\"", start);
        messageId = Utils::StringToWString(response.substr(start, end - start));

        return true;
    }

    bool AddSmallAttachment(const std::wstring& messageId, const std::wstring& filePath)
    {
        std::vector<BYTE> fileData = Utils::ReadFileA(filePath);
        if (fileData.empty()) return false;

        std::wstring fileName = Utils::GetFileName(filePath);
        std::wstring mimeType = Utils::GetMimeType(fileName);
        std::wstring base64Data = Utils::Base64Encode(fileData);

        HINTERNET hSession = WinHttpOpen(L"EmailClient/1.0", WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
            WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0);
        if (hSession == NULL) return false;

        HINTERNET hConnect = WinHttpConnect(hSession, L"graph.microsoft.com",
            INTERNET_DEFAULT_HTTPS_PORT, 0);
        if (hConnect == NULL)
        {
            WinHttpCloseHandle(hSession);
            return false;
        }

        std::wstring path = L"/v1.0/users/" + m_userId + L"/messages/" + messageId + L"/attachments";
        HINTERNET hRequest = WinHttpOpenRequest(hConnect, L"POST", path.c_str(), NULL,
            WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES,
            WINHTTP_FLAG_SECURE);
        if (hRequest == NULL)
        {
            WinHttpCloseHandle(hConnect);
            WinHttpCloseHandle(hSession);
            return false;
        }

        std::wstring headers = L"Authorization: Bearer " + m_tokenManager.Token() + L"\r\n"
            L"Content-Type: application/json\r\n";
        WinHttpAddRequestHeaders(hRequest, headers.c_str(), (DWORD)headers.length(), WINHTTP_ADDREQ_FLAG_ADD);

        std::wstring json = L"{\"@odata.type\":\"#microsoft.graph.fileAttachment\","
            L"\"name\":\"" + fileName + L"\","
            L"\"contentType\":\"" + mimeType + L"\","
            L"\"contentBytes\":\"" + base64Data + L"\"}";

        if (WinHttpSendRequest(hRequest, NULL, 0, (LPVOID)json.c_str(),
            (DWORD)json.length() * sizeof(wchar_t),
            (DWORD)json.length() * sizeof(wchar_t), 0) == FALSE)
        {
            WinHttpCloseHandle(hRequest);
            WinHttpCloseHandle(hConnect);
            WinHttpCloseHandle(hSession);
            return false;
        }

        WinHttpReceiveResponse(hRequest, NULL);
        WinHttpCloseHandle(hRequest);
        WinHttpCloseHandle(hConnect);
        WinHttpCloseHandle(hSession);

        return true;
    }

    bool AddLargeAttachment(const std::wstring& messageId, const std::wstring& filePath)
    {
        std::vector<BYTE> fileData = Utils::ReadFileA(filePath);
        if (fileData.empty()) return false;

        std::wstring fileName = Utils::GetFileName(filePath);
        DWORD fileSize = (DWORD)fileData.size();

        // Step 1: Create upload session
        HINTERNET hSession = WinHttpOpen(L"EmailClient/1.0", WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
            WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0);
        if (hSession == NULL) return false;

        HINTERNET hConnect = WinHttpConnect(hSession, L"graph.microsoft.com",
            INTERNET_DEFAULT_HTTPS_PORT, 0);
        if (hConnect == NULL)
        {
            WinHttpCloseHandle(hSession);
            return false;
        }

        std::wstring path = L"/v1.0/users/" + m_userId + L"/messages/" + messageId + L"/attachments/createUploadSession";
        HINTERNET hRequest = WinHttpOpenRequest(hConnect, L"POST", path.c_str(), NULL,
            WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES,
            WINHTTP_FLAG_SECURE);
        if (hRequest == NULL)
        {
            WinHttpCloseHandle(hConnect);
            WinHttpCloseHandle(hSession);
            return false;
        }

        std::wstring headers = L"Authorization: Bearer " + m_tokenManager.Token() + L"\r\n"
            L"Content-Type: application/json\r\n";
        WinHttpAddRequestHeaders(hRequest, headers.c_str(), (DWORD)headers.length(), WINHTTP_ADDREQ_FLAG_ADD);

        std::wstring json = L"{\"AttachmentItem\":{"
            L"\"attachmentType\":\"file\","
            L"\"name\":\"" + fileName + L"\","
            L"\"size\":" + std::to_wstring((long long)fileSize) + L"}}";

        if (WinHttpSendRequest(hRequest, NULL, 0, (LPVOID)json.c_str(),
            (DWORD)json.length() * sizeof(wchar_t),
            (DWORD)json.length() * sizeof(wchar_t), 0) == FALSE)
        {
            WinHttpCloseHandle(hRequest);
            WinHttpCloseHandle(hConnect);
            WinHttpCloseHandle(hSession);
            return false;
        }

        if (WinHttpReceiveResponse(hRequest, NULL) == FALSE)
        {
            WinHttpCloseHandle(hRequest);
            WinHttpCloseHandle(hConnect);
            WinHttpCloseHandle(hSession);
            return false;
        }

        DWORD size = 0, downloaded = 0;
        std::string response;
        do
        {
            size = 0;
            if (WinHttpQueryDataAvailable(hRequest, &size) == FALSE) break;

            std::vector<char> buffer(size + 1);
            if (WinHttpReadData(hRequest, &buffer[0], size, &downloaded) == FALSE) break;
            response.append(&buffer[0], size);
        } while (size > 0);

        WinHttpCloseHandle(hRequest);

        // Parse upload URL
        size_t urlPos = response.find("uploadUrl");
        if (urlPos == std::string::npos)
        {
            WinHttpCloseHandle(hConnect);
            WinHttpCloseHandle(hSession);
            return false;
        }

        size_t start = response.find("\"", urlPos + 10) + 1;
        size_t end = response.find("\"", start);
        std::wstring uploadUrl = Utils::StringToWString(response.substr(start, end - start));

        // Step 2: Upload in chunks
        DWORD offset = 0;
        DWORD remaining = fileSize;
        bool success = true;

        while (remaining > 0 && success)
        {
            DWORD chunkSize = (remaining > CHUNK_SIZE) ? CHUNK_SIZE : remaining;

            // Extract host and path from URL
            size_t hostStart = uploadUrl.find(L"//") + 2;
            size_t hostEnd = uploadUrl.find(L"/", hostStart);
            std::wstring host = uploadUrl.substr(hostStart, hostEnd - hostStart);
            std::wstring uploadPath = uploadUrl.substr(hostEnd);

            HINTERNET hUploadConnect = WinHttpConnect(hSession, host.c_str(),
                INTERNET_DEFAULT_HTTPS_PORT, 0);
            if (hUploadConnect == NULL)
            {
                success = false;
                break;
            }

            HINTERNET hUploadRequest = WinHttpOpenRequest(hUploadConnect, L"PUT", uploadPath.c_str(),
                NULL, WINHTTP_NO_REFERER,
                WINHTTP_DEFAULT_ACCEPT_TYPES,
                WINHTTP_FLAG_SECURE);
            if (hUploadRequest == NULL)
            {
                WinHttpCloseHandle(hUploadConnect);
                success = false;
                break;
            }

            headers = L"Authorization: Bearer " + m_tokenManager.Token() + L"\r\n"
                L"Content-Type: application/octet-stream\r\n"
                L"Content-Range: bytes " + std::to_wstring((long long)offset) + L"-" +
                std::to_wstring((long long)(offset + chunkSize - 1)) + L"/" +
                std::to_wstring((long long)fileSize) + L"\r\n";
            WinHttpAddRequestHeaders(hUploadRequest, headers.c_str(), (DWORD)headers.length(), WINHTTP_ADDREQ_FLAG_ADD);

            if (WinHttpSendRequest(hUploadRequest, NULL, 0, &fileData[0] + offset,
                chunkSize, chunkSize, 0) == FALSE)
            {
                WinHttpCloseHandle(hUploadRequest);
                WinHttpCloseHandle(hUploadConnect);
                success = false;
                break;
            }

            if (WinHttpReceiveResponse(hUploadRequest, NULL) == FALSE)
            {
                WinHttpCloseHandle(hUploadRequest);
                WinHttpCloseHandle(hUploadConnect);
                success = false;
                break;
            }

            // Read response (required even if we don't use it)
            do
            {
                size = 0;
                if (WinHttpQueryDataAvailable(hUploadRequest, &size) == FALSE) break;

                std::vector<char> buffer(size + 1);
                if (WinHttpReadData(hUploadRequest, &buffer[0], size, &downloaded) == FALSE) break;
            } while (size > 0);

            WinHttpCloseHandle(hUploadRequest);
            WinHttpCloseHandle(hUploadConnect);

            offset += chunkSize;
            remaining -= chunkSize;
        }

        WinHttpCloseHandle(hConnect);
        WinHttpCloseHandle(hSession);
        return success;
    }

    bool SendMessage(const std::wstring& messageId)
    {
        HINTERNET hSession = WinHttpOpen(L"EmailClient/1.0", WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
            WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0);
        if (hSession == NULL) return false;

        HINTERNET hConnect = WinHttpConnect(hSession, L"graph.microsoft.com",
            INTERNET_DEFAULT_HTTPS_PORT, 0);
        if (hConnect == NULL)
        {
            WinHttpCloseHandle(hSession);
            return false;
        }

        std::wstring path = L"/v1.0/users/" + m_userId + L"/messages/" + messageId + L"/send";
        HINTERNET hRequest = WinHttpOpenRequest(hConnect, L"POST", path.c_str(), NULL,
            WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES,
            WINHTTP_FLAG_SECURE);
        if (hRequest == NULL)
        {
            WinHttpCloseHandle(hConnect);
            WinHttpCloseHandle(hSession);
            return false;
        }

        std::wstring headers = L"Authorization: Bearer " + m_tokenManager.Token() + L"\r\n"
            L"Content-Length: 0\r\n";
        WinHttpAddRequestHeaders(hRequest, headers.c_str(), (DWORD)headers.length(), WINHTTP_ADDREQ_FLAG_ADD);

        if (WinHttpSendRequest(hRequest, NULL, 0, WINHTTP_NO_REQUEST_DATA, 0, 0, 0) == FALSE)
        {
            WinHttpCloseHandle(hRequest);
            WinHttpCloseHandle(hConnect);
            WinHttpCloseHandle(hSession);
            return false;
        }

        if (WinHttpReceiveResponse(hRequest, NULL) == FALSE)
        {
            WinHttpCloseHandle(hRequest);
            WinHttpCloseHandle(hConnect);
            WinHttpCloseHandle(hSession);
            return false;
        }

        WinHttpCloseHandle(hRequest);
        WinHttpCloseHandle(hConnect);
        WinHttpCloseHandle(hSession);

        return true;
    }

public:
    EmailSender(const std::wstring& clientId, const std::wstring& clientSecret,
        const std::wstring& tenantId, const std::wstring& userId)
        : m_userId(userId)
    {
        m_tokenManager.GetAccessToken(clientId, clientSecret, tenantId,
            L"https://graph.microsoft.com/.default");
    }

    bool SendEmail(const std::wstring& to, const std::wstring& subject,
        const std::wstring& body, const std::vector<std::wstring>& attachments)
    {
        // Display author information
        wprintf(L"\n📧 Microsoft Graph Email Sender\n");
        wprintf(L"Copyright © 2024 Arulmurugan K\n");
        wprintf(L"Version 1.0.2\n\n");

        // 1. Create draft message
        std::wstring messageId;
        if (!CreateDraftMessage(to, subject, body, messageId))
        {
            wprintf(L"✗ Failed to create draft message\n");
            return false;
        }
        wprintf(L"✓ Draft message created\n");

        // 2. Add all attachments
        int successCount = 0;
        for (size_t i = 0; i < attachments.size(); ++i)
        {
            const std::wstring& filePath = attachments[i];
            wprintf(L"  Processing: %s... ", Utils::GetFileName(filePath).c_str());

            std::vector<BYTE> fileData = Utils::ReadFileA(filePath);
            if (fileData.empty())
            {
                wprintf(L"✗ (file not found)\n");
                continue;
            }

            bool result;
            if (fileData.size() <= MAX_SMALL_FILE_SIZE)
            {
                result = AddSmallAttachment(messageId, filePath);
                wprintf(result ? L"✓ (direct)\n" : L"✗\n");
            }
            else
            {
                result = AddLargeAttachment(messageId, filePath);
                wprintf(result ? L"✓ (chunked)\n" : L"✗\n");
            }

            if (result) successCount++;
        }

        wprintf(L"✓ Attachments: %d of %d added successfully\n", 
                successCount, attachments.size());

        // 3. Send the message
        if (SendMessage(messageId))
        {
            wprintf(L"✓ Email sent successfully to %s\n\n", to.c_str());
            return true;
        }
        else
        {
            wprintf(L"✗ Failed to send email\n\n");
            return false;
        }
    }
};

// Main function
int main()
{
    wprintf(L"╔════════════════════════════════════════════════════════════╗\n");
    wprintf(L"║     Microsoft Graph Email Sender with OAuth 2.0           ║\n");
    wprintf(L"║              Copyright © 2024 Arulmurugan K               ║\n");
    wprintf(L"║                    Version 1.0.2                           ║\n");
    wprintf(L"╚════════════════════════════════════════════════════════════╝\n\n");

    // Configuration - REPLACE WITH YOUR ACTUAL VALUES
    std::wstring clientId = L"your_client_id";
    std::wstring clientSecret = L"your_client_secret";
    std::wstring tenantId = L"your_tenant_id";
    std::wstring userId = L"user@domain.com"; // or "me" for current user

    // Email details
    std::wstring to = L"recipient@example.com";
    std::wstring subject = L"Test Email with Multiple Attachments";
    std::wstring body = L"This email contains multiple attachments of different sizes.";

    // Attachments - add your actual file paths
    std::vector<std::wstring> attachments;
    attachments.push_back(L"C:\\path\\to\\small_file1.txt");    // <3MB
    attachments.push_back(L"C:\\path\\to\\small_file2.jpg");    // <3MB
    attachments.push_back(L"C:\\path\\to\\large_file1.pdf");    // >3MB
    attachments.push_back(L"C:\\path\\to\\large_file2.zip");    // >3MB

    // Send email
    EmailSender sender(clientId, clientSecret, tenantId, userId);
    if (sender.SendEmail(to, subject, body, attachments))
    {
        wprintf(L"✅ Email with multiple attachments sent successfully!\n");
    }
    else
    {
        wprintf(L"❌ Failed to send email with attachments.\n");
        wprintf(L"   Please check your configuration and try again.\n");
    }

    wprintf(L"\n📝 This software is provided under MIT License\n");
    wprintf(L"   Copyright (c) 2024 Arulmurugan K. All rights reserved.\n");
    wprintf(L"   GitHub: https://github.com/arulmurugank/graph-email-sender\n\n");

    system("pause");
    return 0;
}

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

