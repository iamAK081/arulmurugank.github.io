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
