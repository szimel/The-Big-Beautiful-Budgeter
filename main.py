import pymupdf

doc = pymupdf.open("./statements/Credit Statement May 2026.pdf")
text = ""

for page in doc:
    text += page.get_text() + "\n"

print(text)