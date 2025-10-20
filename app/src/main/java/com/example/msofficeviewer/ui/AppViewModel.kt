package com.example.msofficeviewer.ui

import android.content.ContentResolver
import android.net.Uri
import androidx.lifecycle.ViewModel
import androidx.lifecycle.viewModelScope
import com.example.msofficeviewer.core.OfficeParser
import com.example.msofficeviewer.core.OfficeType
import com.example.msofficeviewer.core.OfficeWriter
import com.example.msofficeviewer.core.toOfficeType
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.flow.MutableStateFlow
import kotlinx.coroutines.flow.StateFlow
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext

data class UiState(
    val isLoading: Boolean = false,
    val mimeType: String = "",
    val uri: Uri? = null,
    val content: String? = null,
    val error: String? = null,
)

class AppViewModel : ViewModel() {
    private val _uiState = MutableStateFlow(UiState())
    val uiState: StateFlow<UiState> = _uiState

    fun openUri(contentResolver: ContentResolver, uri: Uri) {
        _uiState.value = UiState(isLoading = true, uri = uri)
        viewModelScope.launch {
            val mime = contentResolver.getType(uri) ?: ""
            val officeType = mime.toOfficeType()
            val result = withContext(Dispatchers.IO) {
                runCatching { OfficeParser.parse(contentResolver.openInputStream(uri)!!, officeType) }
            }
            _uiState.value = result.fold(
                onSuccess = { UiState(isLoading = false, mimeType = mime, uri = uri, content = it, error = null) },
                onFailure = { UiState(isLoading = false, mimeType = mime, uri = uri, content = null, error = it.localizedMessage ?: it.javaClass.simpleName) }
            )
        }
    }

    fun updateContent(newContent: String) {
        _uiState.value = _uiState.value.copy(content = newContent)
    }

    suspend fun saveDocx(contentResolver: ContentResolver) {
        val current = _uiState.value
        if (current.content == null) return
        val bytes = withContext(Dispatchers.Default) {
            OfficeWriter.writeDocxFromPlainText(current.content)
        }
        // Let user choose location via CreateDocument
        // In a production app, wire this via ActivityResult, but here we overwrite the same URI if possible.
        val uri = current.uri ?: return
        withContext(Dispatchers.IO) {
            contentResolver.openOutputStream(uri, "rwt")?.use { it.write(bytes) }
        }
    }

    fun clear() {
        _uiState.value = UiState()
    }

    fun retry(contentResolver: ContentResolver) {
        val last = _uiState.value.uri ?: return
        openUri(contentResolver, last)
    }
}


