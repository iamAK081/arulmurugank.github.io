---
layout: post
title: "Learning C++ Pointers"
date: 2024-03-20
author: Arulmurugan K
categories: [C++, Tutorial]
---

# Understanding C++ Pointers

Pointers are variables that store memory addresses...

## Code Example
```cpp
int x = 10;
int* ptr = &x;


## 🎯 Which Method Should You Use?

| Your Skill Level | Recommended Method |
|-----------------|-------------------|
| **Beginner** | Method 2 (Multiple static pages) |
| **Intermediate** | Method 1 (Edit post.html) |
| **Advanced** | Method 3 (Jekyll posts) |

## ✅ Simple Recommendation for You

Since you're just starting, let's keep it **SUPER SIMPLE**:

### Use Method 2: Create Separate Pages

1. **For your current article**, you have `post.html` - keep it
2. **For new article**, create `article2.html`
3. **Update `index.html`** to link both

Your `index.html` would look like:
```html
<h1>Arulmurugan K's Blog</h1>

<div class="post">
    <h2>Microsoft Graph Email Sender</h2>
    <p>February 28, 2024</p>
    <a href="post.html">Read More</a>
</div>

<div class="post">
    <h2>My New Article Title</h2>
    <p>March 20, 2024</p>
    <a href="article2.html">Read More</a>
</div>
