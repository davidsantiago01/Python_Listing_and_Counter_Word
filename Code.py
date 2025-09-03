

from collections import Counter
from docx import Document
import re

# Stopwords you want to exclude. These are in Spanish, modify at will
STOPWORDS = {"es", "que", "el", "una", "un", "la", "lo", "y", "me", "yo",
             "he", "en", "de", "los", "las", "con", "sin", "a", "se", "por"
             "para", "mi", "te", "tu", "le", "les", "del", "al", "si", "no",
             "su", "sus", "este", "esta", "estos", "estas", "eso", "esa",
             "esos", "esas", "así", "como", "más", "pero", "o", "u", "fue",
             "por", "mis", "nos", "ese", "aquí", "ya", "tus", "muy", "para", "algo"
               "algún"  }

def read_docx(file_path):
    """Reads text from a Word (.docx) file."""
    doc = Document(file_path)
    text = []
    for para in doc.paragraphs:
        text.append(para.text)
    return "\n".join(text)

def get_top_words(text, top_n=50):
    """Returns the top N most frequent words in the text (excluding stopwords)."""
  # Normalize: lowercase, remove punctuation, split into words. \b servers as word boundary to match complete words, [a-záéíóúüñ] we including lower case English letter but also special characters from Spanish language, sicne my text is in Spanish.
    words = re.findall(r'\b[a-záéíóúüñ]+\b', text.lower())
    
   
    filtered_words = [w for w in words if w not in STOPWORDS]
    
    
    counter = Counter(filtered_words)
    
    return counter.most_common(top_n)

if __name__ == "__main__":
    file_path = "File name or file path"  #Place your file where Python script.  If is in another location, provide full path 
    text = read_docx(file_path)
    
    top_words = get_top_words(text, 50)
    
    print("Top 50 Most Used Words (excluding stopwords):\n")
    for word, count in top_words:
        print(f"{word}: {count}")
