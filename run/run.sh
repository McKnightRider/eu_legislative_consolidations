#!/usr/bin/env bash
set -euo pipefail

python stage1.py inputs/260706_EUPR_Original.html outputs/stage1.docx &&
python stage2.py --html inputs/260706_EUPR_Original.html --pdf inputs/260706_EUPR_Original.pdf outputs/stage1.docx &&

python stage3.py outputs/stage1.docx outputs/stage3_1.docx inputs/260706_EUPR_FirstAmendment.html orange &&
python stage4.py --html consolidated/260711_Consolidated_FirstAmendment.html --pdf consolidated/260711_Consolidated_FirstAmendment.pdf outputs/stage3_1.docx &&

python stage3.py outputs/stage3_1.docx outputs/stage3_2.docx inputs/260706_EUPR_SecondAmendment.html purple &&
python stage4.py --html consolidated/260711_Consolidated_SecondAmendment.html --pdf consolidated/260711_Consolidated_SecondAmendment.pdf outputs/stage3_2.docx &&

python stage3.py outputs/stage3_2.docx outputs/stage3_3.docx inputs/260706_EUPR_ThirdAmendment.html blue &&
python stage4.py --html consolidated/260711_Consolidated_ThirdAmendment.html --pdf consolidated/260711_Consolidated_ThirdAmendment.pdf outputs/stage3_3.docx &&

python stage3.py outputs/stage3_3.docx outputs/stage3_4.docx inputs/260706_EUPR_FourthAmendment.html brown &&
python stage4.py --html consolidated/260711_Consolidated_FourthAmendment.html --pdf consolidated/260711_Consolidated_FourthAmendment.pdf outputs/stage3_4.docx &&

python stage3.py outputs/stage3_4.docx outputs/stage3_5.docx inputs/260706_EUPR_FifthAmendment.html green &&
python stage4.py --html consolidated/260711_Consolidated_FifthAmendment.html --pdf consolidated/260711_Consolidated_FifthAmendment.pdf outputs/stage3_5.docx
