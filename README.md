# Word-PDF-Document Translation    
the repository serves the purpose to translate document in Word format from the source language to the target language.    
it makes use of Microsoft Translation API (text translation), to translate Word document and remains its document styles unchanged. Even though Microsoft word has innerly embeded the word document translation feature, there remains have advantages in two point:
- 1. it supports dynamic dictionary, that could help literatly translate some specialized vocabulary, specialized Name, etc, into known words, even though dynamic dictionary holds its limitation which Microsoft Translation doesn't advice to often use.
- 1. it exacts the paragraph from page, rather than phrases or pieces of pharagraph before sending to Microsoft Translation API to generate the translated text. the Python-docx library defines the conception of "Run" to divide the paragraph into pieces of pharagraph or phrases in order to differentiate the font styles. Such would let the Microsoft Document Translation feature break the pharagraph into pieces of words, or phrases, which lead the paragraph unnatually.
