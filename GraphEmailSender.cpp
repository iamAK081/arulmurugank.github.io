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