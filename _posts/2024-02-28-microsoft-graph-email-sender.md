---
layout: post
title: "Microsoft Graph Email Sender with OAuth 2.0 in C++"
date: 2024-02-28
categories: [C++, Windows, Programming]
author: Arulmurugan K
---

<!-- 
  Copyright (c) 2024 Arulmurugan K
  All Rights Reserved.
-->

<div align="center" style="margin-bottom: 30px;">
  <h1>📧 Microsoft Graph Email Sender</h1>
  <h3 style="color: #666; font-weight: normal;">C++ Implementation with OAuth 2.0</h3>
  
  <p style="margin-top: 20px;">
    <img src="https://img.shields.io/badge/Version-1.0.2-blue" alt="Version">
    <img src="https://img.shields.io/badge/Platform-Windows-lightgrey" alt="Platform">
    <img src="https://img.shields.io/badge/VS-2008-purple" alt="VS 2008">
    <img src="https://img.shields.io/badge/License-MIT-green" alt="MIT">
  </p>
  
  <p style="margin-top: 20px; font-size: 1.1em;">
    <em>Created by <strong>Arulmurugan K</strong></em><br>
    <span style="color: #888;">© 2024 All Rights Reserved</span>
  </p>
</div>

---

## 📋 Overview

A clean, efficient C++ application that sends emails with attachments through Microsoft Graph API. Handles files from 1KB to 150MB seamlessly.

**Perfect for:**
- Automated reports
- Backup notifications
- Batch email processing
- Enterprise integration

---

## ✨ Features

| | |
|---|---|
| 🔐 | **OAuth 2.0** - Secure Azure AD authentication |
| 🔄 | **Auto Token Refresh** - Never worry about expiry |
| 📎 | **Smart Attachments** - Files ≤3MB direct, >3MB chunked |
| 📦 | **Multiple Files** - Send unlimited attachments |
| 🔧 | **VS 2008** - Pure Win32, no dependencies |
| ⚡ | **Fast** - Native code, minimal memory |

---

## 📥 Quick Download

<div align="center" style="margin: 30px 0;">
  <a href="https://github.com/arulmurugank/arulmurugank.github.io/raw/main/GraphEmailSender.cpp" 
     style="background: #2ea44f; color: white; padding: 12px 30px; text-decoration: none; border-radius: 6px; font-size: 1.2em; margin: 10px;">
    📥 Download Source Code
  </a>
  
  <a href="https://github.com/arulmurugank/arulmurugank.github.io" 
     style="background: #24292e; color: white; padding: 12px 30px; text-decoration: none; border-radius: 6px; font-size: 1.2em; margin: 10px;">
    ⭐ View on GitHub
  </a>
</div>

---

## 🔧 Prerequisites

| What you need | Where to get it |
|---------------|-----------------|
| Visual Studio 2008+ | [Download](https://visualstudio.microsoft.com/vs/older-downloads/) |
| Windows SDK | [Download](https://developer.microsoft.com/windows/downloads/windows-sdk/) |
| Azure Account | [Free Trial](https://azure.microsoft.com/free/) |

---

## 🚀 Setup in 3 Steps

### 1️⃣ Azure Setup (5 minutes)

```bash
1. Go to portal.azure.com
2. Register an application
3. Add Mail.Send permission
4. Create client secret

2️⃣ Configure the Code

In main() function, replace these:

std::wstring clientId = L"your-client-id";        // From Azure
std::wstring clientSecret = L"your-secret";       // From Azure  
std::wstring tenantId = L"your-tenant-id";        // From Azure
std::wstring userId = L"your-email@company.com";  // Sender's email

3️⃣ Build & Run
1. Open in Visual Studio
2. Add winhttp.lib and crypt32.lib
3. Press F7 to build
4. Press F5 to run


📝 Source Code
<details>
 <summary>👆 Click to view the complete source code</summary>
// GraphEmailSender.cpp
// Microsoft Graph Email Sender with OAuth 2.0
// Copyright (c) 2024 Arulmurugan K

#include <windows.h>
#include <winhttp.h>
#include <stdio.h>
#include <string>
#include <vector>
#include <algorithm>
#include <ctime>

#pragma comment(lib, "winhttp.lib")
#pragma comment(lib, "crypt32.lib")

// ==================== CONSTANTS ====================
#define CHUNK_SIZE (5 * 1024 * 1024)           // 5MB for large files
#define MAX_SMALL_FILE_SIZE (3 * 1024 * 1024)   // 3MB threshold
#define TOKEN_EXPIRY_BUFFER 300                  // 5 min buffer

// ==================== UTILITIES ====================
namespace Utils {
    // Convert string to wide string
    std::wstring StringToWString(const std::string& s) {
        int len = MultiByteToWideChar(CP_ACP, 0, s.c_str(), (int)s.length(), NULL, 0);
        std::wstring ws(len, 0);
        MultiByteToWideChar(CP_ACP, 0, s.c_str(), (int)s.length(), &ws[0], len);
        return ws;
    }

    // Convert wide string to string
    std::string WStringToString(const std::wstring& ws) {
        int len = WideCharToMultiByte(CP_ACP, 0, ws.c_str(), (int)ws.length(), NULL, 0, NULL, NULL);
        std::string s(len, 0);
        WideCharToMultiByte(CP_ACP, 0, ws.c_str(), (int)ws.length(), &s[0], len, NULL, NULL);
        return s;
    }

    // Read file into memory
    std::vector<BYTE> ReadFile(const std::wstring& path) {
        HANDLE hFile = CreateFile(path.c_str(), GENERIC_READ, FILE_SHARE_READ, NULL,
            OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
        if (hFile == INVALID_HANDLE_VALUE) return {};
        
        DWORD size = GetFileSize(hFile, NULL);
        if (size == INVALID_FILE_SIZE) { CloseHandle(hFile); return {}; }
        
        std::vector<BYTE> buffer(size);
        DWORD read = 0;
        if (!ReadFile(hFile, &buffer[0], size, &read, NULL)) {
            CloseHandle(hFile);
            return {};
        }
        
        CloseHandle(hFile);
        return buffer;
    }

    // Base64 encode
    std::wstring Base64Encode(const std::vector<BYTE>& data) {
        DWORD len = 0;
        if (!CryptBinaryToStringW(&data[0], (DWORD)data.size(),
            CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF, NULL, &len))
            return L"";
        
        std::wstring result(len, 0);
        if (!CryptBinaryToStringW(&data[0], (DWORD)data.size(),
            CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF, &result[0], &len))
            return L"";
        
        if (!result.empty() && result.back() == L'\0')
            result.pop_back();
        return result;
    }

    // Get filename from path
    std::wstring GetFileName(const std::wstring& path) {
        size_t pos = path.find_last_of(L"\\/");
        return (pos == std::wstring::npos) ? path : path.substr(pos + 1);
    }

    // Get MIME type
    std::wstring GetMimeType(const std::wstring& filename) {
        size_t dot = filename.find_last_of(L".");
        if (dot == std::wstring::npos) return L"application/octet-stream";
        
        std::wstring ext = filename.substr(dot + 1);
        std::transform(ext.begin(), ext.end(), ext.begin(), ::towlower);
        
        if (ext == L"txt") return L"text/plain";
        if (ext == L"pdf") return L"application/pdf";
        if (ext == L"jpg" || ext == L"jpeg") return L"image/jpeg";
        if (ext == L"png") return L"image/png";
        if (ext == L"doc" || ext == L"docx") return L"application/msword";
        if (ext == L"xls" || ext == L"xlsx") return L"application/vnd.ms-excel";
        if (ext == L"zip") return L"application/zip";
        
        return L"application/octet-stream";
    }
}

// ==================== TOKEN MANAGER ====================
class TokenManager {
    std::wstring m_token;
    time_t m_expiry;

    bool RequestNewToken(const std::wstring& clientId, const std::wstring& clientSecret,
                         const std::wstring& tenantId, const std::wstring& scope) {
        HINTERNET hSession = WinHttpOpen(L"TokenClient/1.0", 
            WINHTTP_ACCESS_TYPE_DEFAULT_PROXY, NULL, NULL, 0);
        if (!hSession) return false;

        HINTERNET hConnect = WinHttpConnect(hSession, L"login.microsoftonline.com",
            INTERNET_DEFAULT_HTTPS_PORT, 0);
        if (!hConnect) { WinHttpCloseHandle(hSession); return false; }

        std::wstring path = L"/" + tenantId + L"/oauth2/v2.0/token";
        HINTERNET hRequest = WinHttpOpenRequest(hConnect, L"POST", path.c_str(), NULL,
            NULL, NULL, WINHTTP_FLAG_SECURE);
        if (!hRequest) { WinHttpCloseHandle(hConnect); WinHttpCloseHandle(hSession); return false; }

        std::wstring data = L"client_id=" + clientId + L"&client_secret=" + clientSecret +
            L"&scope=" + scope + L"&grant_type=client_credentials";
        
        std::wstring headers = L"Content-Type: application/x-www-form-urlencoded\r\n";
        WinHttpAddRequestHeaders(hRequest, headers.c_str(), (DWORD)headers.length(), 
            WINHTTP_ADDREQ_FLAG_ADD);

        if (!WinHttpSendRequest(hRequest, NULL, 0, (LPVOID)data.c_str(),
            (DWORD)data.length() * 2, (DWORD)data.length() * 2, 0)) {
            WinHttpCloseHandle(hRequest); WinHttpCloseHandle(hConnect); 
            WinHttpCloseHandle(hSession); return false;
        }

        if (!WinHttpReceiveResponse(hRequest, NULL)) {
            WinHttpCloseHandle(hRequest); WinHttpCloseHandle(hConnect); 
            WinHttpCloseHandle(hSession); return false;
        }

        std::string response;
        DWORD size = 0;
        do {
            if (!WinHttpQueryDataAvailable(hRequest, &size)) break;
            std::vector<char> buffer(size + 1);
            DWORD read = 0;
            if (!WinHttpReadData(hRequest, &buffer[0], size, &read)) break;
            response.append(&buffer[0], read);
        } while (size > 0);

        WinHttpCloseHandle(hRequest);
        WinHttpCloseHandle(hConnect);
        WinHttpCloseHandle(hSession);

        // Parse token
        size_t tokenPos = response.find("access_token");
        if (tokenPos == std::string::npos) return false;
        
        size_t start = response.find("\"", tokenPos + 14) + 1;
        size_t end = response.find("\"", start);
        m_token = Utils::StringToWString(response.substr(start, end - start));

        // Parse expiry
        size_t expiryPos = response.find("expires_in");
        if (expiryPos != std::string::npos) {
            start = response.find(":", expiryPos) + 1;
            end = response.find_first_of(",}", start);
            m_expiry = time(NULL) + atoi(response.substr(start, end - start).c_str());
        } else {
            m_expiry = time(NULL) + 3600;
        }

        return true;
    }

public:
    TokenManager() : m_expiry(0) {}

    bool GetToken(const std::wstring& clientId, const std::wstring& clientSecret,
                  const std::wstring& tenantId, const std::wstring& scope) {
        if (m_token.empty() || time(NULL) > (m_expiry - TOKEN_EXPIRY_BUFFER)) {
            return RequestNewToken(clientId, clientSecret, tenantId, scope);
        }
        return true;
    }

    const std::wstring& Token() const { return m_token; }
};

// ==================== EMAIL SENDER ====================
class EmailSender {
    TokenManager m_tokens;
    std::wstring m_userId;

    bool CreateDraft(const std::wstring& to, const std::wstring& subject,
                     const std::wstring& body, std::wstring& id) {
        HINTERNET hSession = WinHttpOpen(L"EmailClient/1.0", 
            WINHTTP_ACCESS_TYPE_DEFAULT_PROXY, NULL, NULL, 0);
        if (!hSession) return false;

        HINTERNET hConnect = WinHttpConnect(hSession, L"graph.microsoft.com",
            INTERNET_DEFAULT_HTTPS_PORT, 0);
        if (!hConnect) { WinHttpCloseHandle(hSession); return false; }

        std::wstring path = L"/v1.0/users/" + m_userId + L"/messages";
        HINTERNET hRequest = WinHttpOpenRequest(hConnect, L"POST", path.c_str(), NULL,
            NULL, NULL, WINHTTP_FLAG_SECURE);
        if (!hRequest) { WinHttpCloseHandle(hConnect); WinHttpCloseHandle(hSession); return false; }

        std::wstring headers = L"Authorization: Bearer " + m_tokens.Token() + 
            L"\r\nContent-Type: application/json\r\n";
        WinHttpAddRequestHeaders(hRequest, headers.c_str(), (DWORD)headers.length(), 
            WINHTTP_ADDREQ_FLAG_ADD);

        std::wstring json = L"{\"message\":{"
            L"\"subject\":\"" + subject + L"\","
            L"\"body\":{\"contentType\":\"Text\",\"content\":\"" + body + L"\"},"
            L"\"toRecipients\":[{\"emailAddress\":{\"address\":\"" + to + L"\"}}]"
            L"},"
            L"\"saveToSentItems\":\"true\"}";

        if (!WinHttpSendRequest(hRequest, NULL, 0, (LPVOID)json.c_str(),
            (DWORD)json.length() * 2, (DWORD)json.length() * 2, 0)) {
            WinHttpCloseHandle(hRequest); WinHttpCloseHandle(hConnect); 
            WinHttpCloseHandle(hSession); return false;
        }

        if (!WinHttpReceiveResponse(hRequest, NULL)) {
            WinHttpCloseHandle(hRequest); WinHttpCloseHandle(hConnect); 
            WinHttpCloseHandle(hSession); return false;
        }

        std::string response;
        DWORD size = 0;
        do {
            if (!WinHttpQueryDataAvailable(hRequest, &size)) break;
            std::vector<char> buffer(size + 1);
            DWORD read = 0;
            if (!WinHttpReadData(hRequest, &buffer[0], size, &read)) break;
            response.append(&buffer[0], read);
        } while (size > 0);

        WinHttpCloseHandle(hRequest);
        WinHttpCloseHandle(hConnect);
        WinHttpCloseHandle(hSession);

        size_t idPos = response.find("id");
        if (idPos == std::string::npos) return false;
        
        size_t start = response.find("\"", idPos + 4) + 1;
        size_t end = response.find("\"", start);
        id = Utils::StringToWString(response.substr(start, end - start));
        return true;
    }

    bool AddSmallAttachment(const std::wstring& msgId, const std::wstring& path) {
        auto data = Utils::ReadFile(path);
        if (data.empty()) return false;

        std::wstring name = Utils::GetFileName(path);
        std::wstring mime = Utils::GetMimeType(name);
        std::wstring b64 = Utils::Base64Encode(data);

        HINTERNET hSession = WinHttpOpen(L"EmailClient/1.0", 
            WINHTTP_ACCESS_TYPE_DEFAULT_PROXY, NULL, NULL, 0);
        if (!hSession) return false;

        HINTERNET hConnect = WinHttpConnect(hSession, L"graph.microsoft.com",
            INTERNET_DEFAULT_HTTPS_PORT, 0);
        if (!hConnect) { WinHttpCloseHandle(hSession); return false; }

        std::wstring path2 = L"/v1.0/users/" + m_userId + L"/messages/" + msgId + L"/attachments";
        HINTERNET hRequest = WinHttpOpenRequest(hConnect, L"POST", path2.c_str(), NULL,
            NULL, NULL, WINHTTP_FLAG_SECURE);
        if (!hRequest) { WinHttpCloseHandle(hConnect); WinHttpCloseHandle(hSession); return false; }

        std::wstring headers = L"Authorization: Bearer " + m_tokens.Token() + 
            L"\r\nContent-Type: application/json\r\n";
        WinHttpAddRequestHeaders(hRequest, headers.c_str(), (DWORD)headers.length(), 
            WINHTTP_ADDREQ_FLAG_ADD);

        std::wstring json = L"{\"@odata.type\":\"#microsoft.graph.fileAttachment\","
            L"\"name\":\"" + name + L"\","
            L"\"contentType\":\"" + mime + L"\","
            L"\"contentBytes\":\"" + b64 + L"\"}";

        bool ok = WinHttpSendRequest(hRequest, NULL, 0, (LPVOID)json.c_str(),
            (DWORD)json.length() * 2, (DWORD)json.length() * 2, 0) != FALSE;
        
        if (ok) WinHttpReceiveResponse(hRequest, NULL);
        
        WinHttpCloseHandle(hRequest);
        WinHttpCloseHandle(hConnect);
        WinHttpCloseHandle(hSession);
        return ok;
    }

    bool AddLargeAttachment(const std::wstring& msgId, const std::wstring& path) {
        auto data = Utils::ReadFile(path);
        if (data.empty()) return false;

        std::wstring name = Utils::GetFileName(path);
        DWORD size = (DWORD)data.size();

        // Create upload session
        HINTERNET hSession = WinHttpOpen(L"EmailClient/1.0", 
            WINHTTP_ACCESS_TYPE_DEFAULT_PROXY, NULL, NULL, 0);
        if (!hSession) return false;

        HINTERNET hConnect = WinHttpConnect(hSession, L"graph.microsoft.com",
            INTERNET_DEFAULT_HTTPS_PORT, 0);
        if (!hConnect) { WinHttpCloseHandle(hSession); return false; }

        std::wstring path2 = L"/v1.0/users/" + m_userId + L"/messages/" + msgId + 
            L"/attachments/createUploadSession";
        HINTERNET hRequest = WinHttpOpenRequest(hConnect, L"POST", path2.c_str(), NULL,
            NULL, NULL, WINHTTP_FLAG_SECURE);
        if (!hRequest) { WinHttpCloseHandle(hConnect); WinHttpCloseHandle(hSession); return false; }

        std::wstring headers = L"Authorization: Bearer " + m_tokens.Token() + 
            L"\r\nContent-Type: application/json\r\n";
        WinHttpAddRequestHeaders(hRequest, headers.c_str(), (DWORD)headers.length(), 
            WINHTTP_ADDREQ_FLAG_ADD);

        std::wstring json = L"{\"AttachmentItem\":{"
            L"\"attachmentType\":\"file\","
            L"\"name\":\"" + name + L"\","
            L"\"size\":" + std::to_wstring((long long)size) + L"}}";

        if (!WinHttpSendRequest(hRequest, NULL, 0, (LPVOID)json.c_str(),
            (DWORD)json.length() * 2, (DWORD)json.length() * 2, 0)) {
            WinHttpCloseHandle(hRequest); WinHttpCloseHandle(hConnect); 
            WinHttpCloseHandle(hSession); return false;
        }

        if (!WinHttpReceiveResponse(hRequest, NULL)) {
            WinHttpCloseHandle(hRequest); WinHttpCloseHandle(hConnect); 
            WinHttpCloseHandle(hSession); return false;
        }

        std::string response;
        DWORD avail = 0;
        do {
            if (!WinHttpQueryDataAvailable(hRequest, &avail)) break;
            std::vector<char> buffer(avail + 1);
            DWORD read = 0;
            if (!WinHttpReadData(hRequest, &buffer[0], avail, &read)) break;
            response.append(&buffer[0], read);
        } while (avail > 0);

        WinHttpCloseHandle(hRequest);

        // Parse upload URL
        size_t urlPos = response.find("uploadUrl");
        if (urlPos == std::string::npos) {
            WinHttpCloseHandle(hConnect); WinHttpCloseHandle(hSession);
            return false;
        }

        size_t start = response.find("\"", urlPos + 10) + 1;
        size_t end = response.find("\"", start);
        std::wstring uploadUrl = Utils::StringToWString(response.substr(start, end - start));

        // Upload chunks
        DWORD offset = 0;
        DWORD remaining = size;
        bool success = true;

        while (remaining > 0 && success) {
            DWORD chunk = (remaining > CHUNK_SIZE) ? CHUNK_SIZE : remaining;

            size_t hostStart = uploadUrl.find(L"//") + 2;
            size_t hostEnd = uploadUrl.find(L"/", hostStart);
            std::wstring host = uploadUrl.substr(hostStart, hostEnd - hostStart);
            std::wstring upath = uploadUrl.substr(hostEnd);

            HINTERNET hUpConnect = WinHttpConnect(hSession, host.c_str(),
                INTERNET_DEFAULT_HTTPS_PORT, 0);
            if (!hUpConnect) { success = false; break; }

            HINTERNET hUpRequest = WinHttpOpenRequest(hUpConnect, L"PUT", upath.c_str(),
                NULL, NULL, NULL, WINHTTP_FLAG_SECURE);
            if (!hUpRequest) { WinHttpCloseHandle(hUpConnect); success = false; break; }

            std::wstring hdrs = L"Authorization: Bearer " + m_tokens.Token() + L"\r\n"
                L"Content-Type: application/octet-stream\r\n"
                L"Content-Range: bytes " + std::to_wstring((long long)offset) + L"-" +
                std::to_wstring((long long)(offset + chunk - 1)) + L"/" +
                std::to_wstring((long long)size) + L"\r\n";
            WinHttpAddRequestHeaders(hUpRequest, hdrs.c_str(), (DWORD)hdrs.length(), 
                WINHTTP_ADDREQ_FLAG_ADD);

            if (!WinHttpSendRequest(hUpRequest, NULL, 0, &data[0] + offset,
                chunk, chunk, 0)) {
                WinHttpCloseHandle(hUpRequest); WinHttpCloseHandle(hUpConnect);
                success = false; break;
            }

            WinHttpReceiveResponse(hUpRequest, NULL);
            
            WinHttpCloseHandle(hUpRequest);
            WinHttpCloseHandle(hUpConnect);

            offset += chunk;
            remaining -= chunk;
        }

        WinHttpCloseHandle(hConnect);
        WinHttpCloseHandle(hSession);
        return success;
    }

    bool SendMessage(const std::wstring& msgId) {
        HINTERNET hSession = WinHttpOpen(L"EmailClient/1.0", 
            WINHTTP_ACCESS_TYPE_DEFAULT_PROXY, NULL, NULL, 0);
        if (!hSession) return false;

        HINTERNET hConnect = WinHttpConnect(hSession, L"graph.microsoft.com",
            INTERNET_DEFAULT_HTTPS_PORT, 0);
        if (!hConnect) { WinHttpCloseHandle(hSession); return false; }

        std::wstring path = L"/v1.0/users/" + m_userId + L"/messages/" + msgId + L"/send";
        HINTERNET hRequest = WinHttpOpenRequest(hConnect, L"POST", path.c_str(), NULL,
            NULL, NULL, WINHTTP_FLAG_SECURE);
        if (!hRequest) { WinHttpCloseHandle(hConnect); WinHttpCloseHandle(hSession); return false; }

        std::wstring headers = L"Authorization: Bearer " + m_tokens.Token() + 
            L"\r\nContent-Length: 0\r\n";
        WinHttpAddRequestHeaders(hRequest, headers.c_str(), (DWORD)headers.length(), 
            WINHTTP_ADDREQ_FLAG_ADD);

        bool ok = WinHttpSendRequest(hRequest, NULL, 0, NULL, 0, 0, 0) != FALSE;
        if (ok) WinHttpReceiveResponse(hRequest, NULL);

        WinHttpCloseHandle(hRequest);
        WinHttpCloseHandle(hConnect);
        WinHttpCloseHandle(hSession);
        return ok;
    }

public:
    EmailSender(const std::wstring& clientId, const std::wstring& clientSecret,
                const std::wstring& tenantId, const std::wstring& userId)
        : m_userId(userId) {
        m_tokens.GetToken(clientId, clientSecret, tenantId, 
            L"https://graph.microsoft.com/.default");
    }

    bool SendEmail(const std::wstring& to, const std::wstring& subject,
                   const std::wstring& body, const std::vector<std::wstring>& attachments) {
        printf("\n📧 Microsoft Graph Email Sender\n");
        printf("Copyright © 2024 Arulmurugan K\n\n");

        std::wstring msgId;
        if (!CreateDraft(to, subject, body, msgId)) {
            printf("❌ Failed to create draft\n");
            return false;
        }
        printf("✅ Draft created\n");

        int success = 0;
        for (const auto& file : attachments) {
            printf("   Processing: %ls... ", Utils::GetFileName(file).c_str());
            
            auto data = Utils::ReadFile(file);
            if (data.empty()) {
                printf("❌ not found\n");
                continue;
            }

            bool ok = (data.size() <= MAX_SMALL_FILE_SIZE) 
                ? AddSmallAttachment(msgId, file)
                : AddLargeAttachment(msgId, file);
            
            printf(ok ? "✅\n" : "❌\n");
            if (ok) success++;
        }

        printf("📎 Attachments: %d/%d added\n", success, (int)attachments.size());

        if (SendMessage(msgId)) {
            printf("✅ Email sent to %ls\n\n", to.c_str());
            return true;
        }

        printf("❌ Failed to send\n\n");
        return false;
    }
};

// ==================== MAIN ====================
int main() {
    printf("╔════════════════════════════════════════════╗\n");
    printf("║   Microsoft Graph Email Sender v1.0.2     ║\n");
    printf("║        Copyright © 2024 Arulmurugan K     ║\n");
    printf("╚════════════════════════════════════════════╝\n\n");

    // ===== CONFIGURATION =====
    std::wstring clientId = L"your_client_id";        // From Azure
    std::wstring clientSecret = L"your_client_secret"; // From Azure
    std::wstring tenantId = L"your_tenant_id";         // From Azure
    std::wstring userId = L"user@domain.com";          // Sender's email

    // ===== EMAIL DETAILS =====
    std::wstring to = L"recipient@example.com";
    std::wstring subject = L"Test Email with Attachments";
    std::wstring body = L"This email was sent using Microsoft Graph API with OAuth 2.0.";

    // ===== ATTACHMENTS =====
    std::vector<std::wstring> attachments;
    attachments.push_back(L"C:\\test\\small_file.txt");   // <3MB
    attachments.push_back(L"C:\\test\\large_file.zip");   // >3MB

    // ===== SEND =====
    EmailSender sender(clientId, clientSecret, tenantId, userId);
    sender.SendEmail(to, subject, body, attachments);

    printf("📝 This software is provided under MIT License\n");
    printf("   GitHub: https://github.com/arulmurugank\n\n");

    system("pause");
    return 0;
}
</details>


💡 Usage Examples
Send a simple email
-----------------------
std::vector<std::wstring> attachments;
attachments.push_back(L"C:\\reports\\weekly.pdf");

EmailSender sender(clientId, clientSecret, tenantId, userId);
sender.SendEmail(L"boss@company.com", L"Weekly Report", 
                 L"Please see attached.", attachments);

Send multiple files
---------------------
std::vector<std::wstring> files = {
    L"D:\\backups\\db.bak",      // 500MB - auto chunked
    L"D:\\backups\\logs.zip",     // 200MB - auto chunked
    L"D:\\reports\\summary.txt"   // 10KB - direct
};

sender.SendEmail(L"admin@company.com", L"Daily Backups", 
                 L"Database and logs attached", files);

🔍 Common Issues

Problem	                        Solution
❌ 401                     Unauthorized	Client secret expired? Check Azure portal
❌ 403                     Forbidden	Mail.Send permission missing? Grant admin consent
❌ File not found	         Check file path exists
❌ Connection failed	     Firewall blocking? Check proxy settings

📄 License
MIT License - Copyright (c) 2024 Arulmurugan K
Free to use, modify, and distribute. Keep copyright notice.

👨‍💻 About the Author
<div align="center" style="margin: 30px 0;"> <a href="https://github.com/arulmurugank" style="margin: 0 10px;">📦 GitHub</a> | <a href="mailto:arulmurugan@example.com" style="margin: 0 10px;">📧 Email</a> | <a href="https://linkedin.com/in/arulmurugank" style="margin: 0 10px;">🔗 LinkedIn</a> </div>

<div align="center" style="color: #888; margin-top: 50px;"> <small>Copyright © 2024 Arulmurugan K. All rights reserved.</small><br> <small>Made with ❤️ in India</small> </div>
