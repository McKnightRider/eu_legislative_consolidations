from bs4 import BeautifulSoup

with open("../html_to_word/260706_EUPR_Original.html", encoding="utf-8") as f:
    soup = BeautifulSoup(f, "lxml")

annex = soup.find(id="anx_I")

for p in annex.find_all(["p", "div"])[:100]:
    print("ID:", p.get("id"))
    print("CLASS:", p.get("class"))
    print("TEXT:", " ".join(p.get_text(" ", strip=True).split()))
    print("-" * 80)


for tag in soup.find(id="anx_I").find_all(True):
    txt = " ".join(tag.get_text(" ", strip=True).split())
    if txt:
        print(tag.name, tag.get("id"), tag.get("class"), txt[:200])
