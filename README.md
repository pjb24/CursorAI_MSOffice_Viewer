# CursorAI_MSOffice_Viewer (Android)

Lightweight Android app built with Kotlin and Jetpack Compose to open, view, and perform basic editing of Office Open XML files (DOCX/XLSX/PPTX) without heavy native Office dependencies.

## Features
- Open files via Android System File Picker (SAF)
- View text from DOCX, XLSX, and PPTX
- Edit plain text of DOCX and Save As a new DOCX
- Handles parsing on background threads

## Limitations
- Formatting, images, tables, styles, shapes, charts are not preserved
- XLSX/PPTX are view-only (text extraction only)
- DOCX editing is plain-text paragraphs only; saves as minimal DOCX
- Large/complex files may not render perfectly

## Tech
- Kotlin, Jetpack Compose, Material 3
- XML parsing and ZIP handling via Android XmlPullParser and Apache Commons Compress
- Simple text extraction via Jsoup

## Build
Open the project in Android Studio. Build and run on a device with Android 7.0+ (API 24+).

## Roadmap
- Richer DOCX editing (styles, lists, tables)
- Better XLSX sheet rendering and formulas
- PPTX slides rendering with layout
- Caching and pagination for large docs
