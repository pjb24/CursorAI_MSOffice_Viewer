package com.example.msofficeviewer.core

import android.util.Xml
import org.apache.commons.compress.archivers.zip.ZipArchiveInputStream
import org.xmlpull.v1.XmlPullParser
import java.io.ByteArrayInputStream
import java.io.InputStream

enum class OfficeType { DOCX, XLSX, PPTX, UNKNOWN }

fun String?.toOfficeType(): OfficeType = when (this) {
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document" -> OfficeType.DOCX
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" -> OfficeType.XLSX
    "application/vnd.openxmlformats-officedocument.presentationml.presentation" -> OfficeType.PPTX
    else -> OfficeType.UNKNOWN
}

object OfficeParser {
    fun parse(input: InputStream, type: OfficeType): String = when (type) {
        OfficeType.DOCX -> parseDocxText(input)
        OfficeType.XLSX -> parseXlsxText(input)
        OfficeType.PPTX -> parsePptxText(input)
        OfficeType.UNKNOWN -> "지원하지 않는 형식입니다."
    }

    private fun parseDocxText(input: InputStream): String {
        ZipArchiveInputStream(input).use { zipIn ->
            var entry = zipIn.nextZipEntry
            while (entry != null) {
                if (!entry.isDirectory && entry.name == "word/document.xml") {
                    return extractTextFromWordDocumentXml(zipIn)
                }
                entry = zipIn.nextZipEntry
            }
        }
        return ""
    }

    private fun extractTextFromWordDocumentXml(input: InputStream): String {
        val parser: XmlPullParser = Xml.newPullParser()
        parser.setInput(input, null)
        val sb = StringBuilder()
        var event = parser.eventType
        while (event != XmlPullParser.END_DOCUMENT) {
            if (event == XmlPullParser.START_TAG && (parser.name == "t" || parser.name == "w:t")) {
                sb.append(parser.nextText())
            }
            if (event == XmlPullParser.START_TAG && (parser.name == "p" || parser.name == "w:p")) {
                if (sb.isNotEmpty()) sb.append('\n')
            }
            event = parser.next()
        }
        return sb.toString()
    }

    private fun parseXlsxText(input: InputStream): String {
        val sharedStringsXml = mutableListOf<ByteArray>()
        val sheetXmls = mutableListOf<ByteArray>()
        ZipArchiveInputStream(input).use { zipIn ->
            var entry = zipIn.nextZipEntry
            while (entry != null) {
                if (!entry.isDirectory) {
                    if (entry.name == "xl/sharedStrings.xml") {
                        sharedStringsXml.add(zipIn.readBytes())
                    } else if (entry.name.startsWith("xl/worksheets/") && entry.name.endsWith(".xml")) {
                        sheetXmls.add(zipIn.readBytes())
                    }
                }
                entry = zipIn.nextZipEntry
            }
        }
        val shared = parseSharedStrings(sharedStringsXml)
        val builder = StringBuilder()
        sheetXmls.forEachIndexed { index, xmlBytes ->
            if (index > 0) builder.append('\n')
            builder.append("Sheet ").append(index + 1).append(':').append('\n')
            builder.append(parseSheet(xmlBytes, shared)).append('\n')
        }
        return builder.toString().trim()
    }

    private fun parseSharedStrings(xmlParts: List<ByteArray>): List<String> {
        if (xmlParts.isEmpty()) return emptyList()
        val parser: XmlPullParser = Xml.newPullParser()
        val strings = ArrayList<String>()
        xmlParts.forEach { bytes ->
            parser.setInput(ByteArrayInputStream(bytes), null)
            var event = parser.eventType
            var inT = false
            val sb = StringBuilder()
            while (event != XmlPullParser.END_DOCUMENT) {
                when (event) {
                    XmlPullParser.START_TAG -> {
                        if (parser.name == "t") {
                            inT = true
                            sb.setLength(0)
                        }
                    }
                    XmlPullParser.TEXT -> if (inT) sb.append(parser.text)
                    XmlPullParser.END_TAG -> {
                        if (parser.name == "t" && inT) {
                            strings.add(sb.toString())
                            inT = false
                        }
                    }
                }
                event = parser.next()
            }
        }
        return strings
    }

    private fun parseSheet(xmlBytes: ByteArray, shared: List<String>): String {
        val parser: XmlPullParser = Xml.newPullParser()
        parser.setInput(ByteArrayInputStream(xmlBytes), null)
        val line = StringBuilder()
        val out = StringBuilder()
        var event = parser.eventType
        var inV = false
        var currentCellType: String? = null
        var pendingValue = StringBuilder()
        fun flushCell() {
            val raw = pendingValue.toString()
            val value = if (currentCellType == "s") {
                raw.toIntOrNull()?.let { idx -> shared.getOrNull(idx) } ?: raw
            } else raw
            if (line.isNotEmpty()) line.append('\t')
            line.append(value)
            pendingValue.setLength(0)
            currentCellType = null
        }
        fun flushRow() {
            if (line.isNotEmpty()) {
                if (out.isNotEmpty()) out.append('\n')
                out.append(line.toString())
                line.setLength(0)
            }
        }
        while (event != XmlPullParser.END_DOCUMENT) {
            when (event) {
                XmlPullParser.START_TAG -> {
                    when (parser.name) {
                        "c" -> currentCellType = parser.getAttributeValue(null, "t")
                        "v" -> { inV = true; pendingValue.setLength(0) }
                    }
                }
                XmlPullParser.TEXT -> if (inV) pendingValue.append(parser.text)
                XmlPullParser.END_TAG -> {
                    when (parser.name) {
                        "v" -> { inV = false }
                        "c" -> flushCell()
                        "row" -> flushRow()
                    }
                }
            }
            event = parser.next()
        }
        flushRow()
        return out.toString()
    }

    private fun parsePptxText(input: InputStream): String {
        val slides = mutableListOf<ByteArray>()
        ZipArchiveInputStream(input).use { zipIn ->
            var entry = zipIn.nextZipEntry
            while (entry != null) {
                if (!entry.isDirectory && entry.name.startsWith("ppt/slides/slide") && entry.name.endsWith(".xml")) {
                    slides.add(zipIn.readBytes())
                }
                entry = zipIn.nextZipEntry
            }
        }
        val builder = StringBuilder()
        slides.forEachIndexed { index, bytes ->
            if (index > 0) builder.append("\n\n")
            builder.append("Slide ").append(index + 1).append(':').append('\n')
            builder.append(extractPptxSlideText(bytes))
        }
        return builder.toString().trim()
    }

    private fun extractPptxSlideText(bytes: ByteArray): String {
        val parser: XmlPullParser = Xml.newPullParser()
        parser.setInput(ByteArrayInputStream(bytes), null)
        val out = StringBuilder()
        var event = parser.eventType
        var inAT = false
        val sb = StringBuilder()
        fun flushAT() {
            if (sb.isNotEmpty()) {
                if (out.isNotEmpty()) out.append(' ')
                out.append(sb.toString())
                sb.setLength(0)
            }
        }
        while (event != XmlPullParser.END_DOCUMENT) {
            when (event) {
                XmlPullParser.START_TAG -> if (parser.name == "a:t") { inAT = true; sb.setLength(0) }
                XmlPullParser.TEXT -> if (inAT) sb.append(parser.text)
                XmlPullParser.END_TAG -> if (parser.name == "a:t") { inAT = false; flushAT() }
            }
            event = parser.next()
        }
        return out.toString()
    }

    private fun readAllText(input: InputStream): String = input.readBytes().toString(Charsets.UTF_8)
}


