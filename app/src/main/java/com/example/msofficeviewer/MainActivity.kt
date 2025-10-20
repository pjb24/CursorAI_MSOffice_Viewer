@file:OptIn(ExperimentalMaterial3Api::class)
package com.example.msofficeviewer

import android.content.Intent
import android.net.Uri
import android.os.Bundle
import androidx.activity.ComponentActivity
import androidx.activity.compose.rememberLauncherForActivityResult
import androidx.activity.compose.setContent
import androidx.activity.result.contract.ActivityResultContracts
import androidx.activity.viewModels
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.rememberScrollState
import androidx.compose.foundation.text.BasicTextField
import androidx.compose.foundation.verticalScroll
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Modifier
import androidx.compose.ui.text.input.TextFieldValue
import androidx.compose.ui.unit.dp
import androidx.core.net.toFile
import com.example.msofficeviewer.core.OfficeParser
import com.example.msofficeviewer.core.OfficeType
import com.example.msofficeviewer.core.toOfficeType
import com.example.msofficeviewer.ui.AppViewModel
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch

class MainActivity : ComponentActivity() {
    private val viewModel: AppViewModel by viewModels()

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)

        val openDoc = registerForActivityResult(ActivityResultContracts.OpenDocument()) { uri: Uri? ->
            uri?.let { viewModel.openUri(contentResolver, it) }
        }

        setContent {
            MaterialTheme {
                val uiState by viewModel.uiState.collectAsState()
                val scope = rememberCoroutineScope()

                // Save-As launcher created inside Compose to access latest content
                Scaffold(
                    topBar = {
                        TopAppBar(title = { Text("MSOffice Viewer") })
                    },
                    floatingActionButton = {
                        ExtendedFloatingActionButton(onClick = {
                            openDoc.launch(arrayOf(
                                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                            ))
                        })
                        { Text("Open") }
                    }
                ) { padding ->
                    Column(Modifier.padding(padding).fillMaxSize().padding(12.dp)) {
                        when {
                            uiState.isLoading -> CircularProgressIndicator()
                            uiState.error != null -> ErrorView(uiState.error!!, onRetry = { viewModel.retry(contentResolver) }, onClear = { viewModel.clear() })
                            uiState.content != null -> {
                                val officeType = uiState.mimeType.toOfficeType()
                                val suggestedName = remember(uiState.uri) {
                                    val defaultName = when (officeType) {
                                        OfficeType.DOCX -> "document.docx"
                                        OfficeType.XLSX -> "workbook.xlsx"
                                        OfficeType.PPTX -> "presentation.pptx"
                                        OfficeType.UNKNOWN -> "document.docx"
                                    }
                                    defaultName
                                }
                                var pendingBytes by remember { mutableStateOf<ByteArray?>(null) }
                                val saveAs = rememberLauncherForActivityResult(ActivityResultContracts.CreateDocument(uiState.mimeType.ifEmpty { "application/vnd.openxmlformats-officedocument.wordprocessingml.document" })) { uri: Uri? ->
                                    val bytes = pendingBytes
                                    if (uri != null && bytes != null) {
                                        contentResolver.openOutputStream(uri, "w")?.use { it.write(bytes) }
                                    }
                                    pendingBytes = null
                                }
                                when (officeType) {
                                    OfficeType.DOCX -> DocxEditor(
                                        text = uiState.content!!,
                                        onTextChange = { viewModel.updateContent(it) },
                                        onSave = {
                                            scope.launch(Dispatchers.Default) {
                                                val bytes = com.example.msofficeviewer.core.OfficeWriter.writeDocxFromPlainText(uiState.content!!)
                                                pendingBytes = bytes
                                                saveAs.launch(suggestedName)
                                            }
                                        }
                                    )
                                    OfficeType.XLSX, OfficeType.PPTX, OfficeType.UNKNOWN -> ReadOnlyView(uiState.content!!)
                                }
                            }
                            else -> Text("파일을 선택하세요 (DOCX/XLSX/PPTX)")
                        }
                    }
                }
            }
        }

        // Handle share/open intents
        intent?.data?.let { data ->
            viewModel.openUri(contentResolver, data)
        }
    }
}

@Composable
private fun ErrorView(message: String, onRetry: () -> Unit, onClear: () -> Unit) {
    Column(Modifier.fillMaxSize(), verticalArrangement = Arrangement.spacedBy(12.dp)) {
        Text(text = "오류: $message")
        Row(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
            Button(onClick = onRetry) { Text("Retry") }
            OutlinedButton(onClick = onClear) { Text("Clear") }
        }
    }
}

@Composable
private fun DocxEditor(text: String, onTextChange: (String) -> Unit, onSave: () -> Unit) {
    var state by remember { mutableStateOf(TextFieldValue(text)) }
    LaunchedEffect(text) { if (text != state.text) state = TextFieldValue(text) }

    Column(Modifier.fillMaxSize()) {
        Row(Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.End) {
            Button(onClick = onSave) { Text("Save") }
        }
        Spacer(Modifier.height(8.dp))
        BasicTextField(
            value = state,
            onValueChange = {
                state = it
                onTextChange(it.text)
            },
            modifier = Modifier.fillMaxSize().verticalScroll(rememberScrollState())
        )
    }
}

@Composable
private fun ReadOnlyView(text: String) {
    Text(text = text, modifier = Modifier.fillMaxSize().verticalScroll(rememberScrollState()))
}


