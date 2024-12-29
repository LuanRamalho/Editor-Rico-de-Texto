import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTextEdit, QToolBar, QAction, QColorDialog, QInputDialog, QFileDialog
)
from PyQt5.QtGui import QIcon, QTextCursor, QFont, QTextCharFormat, QTextListFormat
from PyQt5.QtCore import Qt

from docx import Document

class TextEditor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Editor de Texto Rico")
        self.setGeometry(100, 100, 800, 600)

        # Caixa de texto
        self.textEdit = QTextEdit(self)
        self.setCentralWidget(self.textEdit)

        # Barra de ferramentas
        toolbar = QToolBar("Barra de Ferramentas", self)
        self.addToolBar(toolbar)

        # Botões de formatação
        self.addToolbarAction(toolbar, "Cor", "color.png", self.changeColor)
        self.addToolbarAction(toolbar, "Aumentar Fonte", "font_increase.png", self.increaseFont)
        self.addToolbarAction(toolbar, "Diminuir Fonte", "font_decrease.png", self.decreaseFont)
        self.addToolbarAction(toolbar, "Negrito", "bold.png", self.toggleBold)
        self.addToolbarAction(toolbar, "Itálico", "italic.png", self.toggleItalic)
        self.addToolbarAction(toolbar, "Sublinhado", "underline.png", self.toggleUnderline)
        self.addToolbarAction(toolbar, "Riscado", "strikethrough.png", self.toggleStrikeThrough)
        self.addToolbarAction(toolbar, "Hiperlink", "link.png", self.insertHyperlink)
        self.addToolbarAction(toolbar, "Sobrescrito", "superscript.png", self.toggleSuperscript)
        self.addToolbarAction(toolbar, "Subscrito", "subscript.png", self.toggleSubscript)
        self.addToolbarAction(toolbar, "Alinhar Esquerda", "align_left.png", lambda: self.alignText(Qt.AlignLeft))
        self.addToolbarAction(toolbar, "Alinhar Centro", "align_center.png", lambda: self.alignText(Qt.AlignCenter))
        self.addToolbarAction(toolbar, "Alinhar Direita", "align_right.png", lambda: self.alignText(Qt.AlignRight))
        self.addToolbarAction(toolbar, "Justificar", "justify.png", lambda: self.alignText(Qt.AlignJustify))
        self.addToolbarAction(toolbar, "Lista de Pontos", "bullet_list.png", self.insertBulletList)
        self.addToolbarAction(toolbar, "Lista de Números", "number_list.png", self.insertNumberList)
        self.addToolbarAction(toolbar, "Salvar Arquivo", "save.png", self.saveArchive)

        self.show()

    def addToolbarAction(self, toolbar, name, icon, method):
        action = QAction(QIcon(icon), name, self)
        action.triggered.connect(method)
        toolbar.addAction(action)

    def changeColor(self):
        color = QColorDialog.getColor()
        if color.isValid():
            self.textEdit.setTextColor(color)

    def increaseFont(self):
        font = self.textEdit.currentFont()
        font.setPointSize(font.pointSize() + 1)
        self.textEdit.setCurrentFont(font)

    def decreaseFont(self):
        font = self.textEdit.currentFont()
        font.setPointSize(max(font.pointSize() - 1, 1))
        self.textEdit.setCurrentFont(font)

    def toggleBold(self):
        fmt = self.textEdit.currentCharFormat()
        fmt.setFontWeight(QFont.Bold if fmt.fontWeight() != QFont.Bold else QFont.Normal)
        self.textEdit.mergeCurrentCharFormat(fmt)

    def toggleItalic(self):
        fmt = self.textEdit.currentCharFormat()
        fmt.setFontItalic(not fmt.fontItalic())
        self.textEdit.mergeCurrentCharFormat(fmt)

    def toggleUnderline(self):
        fmt = self.textEdit.currentCharFormat()
        fmt.setFontUnderline(not fmt.fontUnderline())
        self.textEdit.mergeCurrentCharFormat(fmt)

    def toggleStrikeThrough(self):
        fmt = self.textEdit.currentCharFormat()
        fmt.setFontStrikeOut(not fmt.fontStrikeOut())
        self.textEdit.mergeCurrentCharFormat(fmt)

    def insertHyperlink(self):
        url, ok = QInputDialog.getText(self, "Inserir Hiperlink", "URL:")
        if ok and url:
            cursor = self.textEdit.textCursor()
            cursor.insertHtml(f'<a href="{url}">{url}</a>')

    def toggleSuperscript(self):
        fmt = self.textEdit.currentCharFormat()
        fmt.setVerticalAlignment(
            QTextCharFormat.AlignSuperScript
            if fmt.verticalAlignment() != QTextCharFormat.AlignSuperScript
            else QTextCharFormat.AlignNormal
        )
        self.textEdit.mergeCurrentCharFormat(fmt)

    def toggleSubscript(self):
        fmt = self.textEdit.currentCharFormat()
        fmt.setVerticalAlignment(
            QTextCharFormat.AlignSubScript
            if fmt.verticalAlignment() != QTextCharFormat.AlignSubScript
            else QTextCharFormat.AlignNormal
        )
        self.textEdit.mergeCurrentCharFormat(fmt)

    def alignText(self, alignment):
        self.textEdit.setAlignment(alignment)

    def insertBulletList(self):
        cursor = self.textEdit.textCursor()
        # Create a new QTextListFormat object with the desired style
        list_format = QTextListFormat()
        list_format.setStyle(QTextListFormat.ListDisc)  # Use ListDisc for bullet points
        cursor.insertList(list_format)

    def insertNumberList(self):
        cursor = self.textEdit.textCursor()
        # Create a new QTextListFormat object with the desired style
        list_format = QTextListFormat()
        list_format.setStyle(QTextListFormat.ListDecimal)  # Use ListDecimal for numbered lists
        cursor.insertList(list_format)

    def saveArchive(self):
        """Saves the text editor content to a DOCX file."""

        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        filename, _ = QFileDialog.getSaveFileName(self, "Salvar Arquivo", "", "Documento do Word (*.docx)", options=options)

        if filename:
            try:
                with open(filename, 'wb') as f:
                    # **Handle DOCX format limitations:**
                    # Due to limitations of using plain Python for saving DOCX,
                    # we'll provide a fallback to saving as plain text (.txt).
                    # Consider using external libraries (e.g., docx) for full DOCX support.
                    text = self.textEdit.toPlainText()
                    f.write(text.encode('utf-8'))  # Encode for cross-platform compatibility
                    self.statusBar().showMessage("Arquivo salvo com sucesso!")
            except Exception as e:
                self.statusBar().showMessage(f"Erro ao salvar arquivo: {e}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    editor = TextEditor()
    sys.exit(app.exec_())
