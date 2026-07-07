# EU Legislative Consolidations
This repository contains code that (a) converts a piece of EU legislation into Word and (b) shows changes to that EU legislation in different colours.

There are several steps to the process:
1. Python-based conversion - structured XML/HTML (not pdf) -> identify provisions (e.g. recitals, articles, footnotes, annexes) -> Word document.
2. Quality check - checks Word document against original legislation (perhaps the pdf version as that can be used for Litera comparisons).
3. Python-based consolidation engine - identify amending regulation -> add recitals -> parse amending provisions -> apply amendments -> new Word document.
4. Quality check - checks Word document changes (perhaps against consolidated pdf version on Europa website).
5. Legislative consolidation assistant - agent that locates documents -> identifies amendment requirements -> invokes consolidation scripts -> return finished Word document.

There are potentially future developments, such as the agent noting where particular amendments came from and official EU/ESMA commentary.
