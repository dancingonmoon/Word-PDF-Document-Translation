# Word-Document Translation    
the repository serves the purpose to translate documents in Word format from the source language to the target language.    
it makes use of Microsoft Translation API (text translation), to translate Word document and remains its document styles unchanged. Even though Microsoft word has innerly embeded the word document translation feature, there remains have advantages in two points:
1. it supports dynamic dictionary, that could help literally translate some specialized vocabulary, specialized Name, etc, into known words or phrases, even though dynamic dictionary holds its limitation which Microsoft Translation even doesn't advice to often use.
1. it exacts the paragraph from page, rather than phrases or pieces of pharagraph,before sending to Microsoft Translation API to generate the translated text. Equally, it is tranlated in paragraph rather than phrases or pieces of words. the Python-docx library defines the conception of "Run" to divide the paragraph into pieces of pharagraph or phrases in order to differentiate the font styles. Such would let the Microsoft Document Translation feature break the pharagraph into pieces of words, or phrases, which lead to less natual translation.
# PDF-Document Translation
To translate PDF document, here it uses pymupdf libary, which doesn't support the text editing, so, to produce a new PDF is the solution, which get all of image copied, and get text in each span/line/block translated and insert back into new PDF with the same distribution the span text and the block text. similarly, it supports:    
1. dynamic dictionary, which you could phrase and translated phrase into dictionary;
2. the Chinese/Japanese/Korea language translation doesn't easily support the word position matches with translated word position, therefore, the code divides the translated block into sequence of translated span via the source span text distribution. sometimes, it divides not exactly as the nature language of the source text, but the character length remains the same distribution.
# Usage:
> import the word_translate and PDF_translate into your code, and the function has its explanation in each of its augment and output.
> 
