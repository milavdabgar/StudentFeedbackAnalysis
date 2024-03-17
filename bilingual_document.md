---
title: "My Bilingual Document"
author: "John Doe"
date: "2023-06-13"
output:
  pdf_document:
    latex_engine: xelatex
header-includes:
  - \usepackage{fontspec}
  - \usepackage{polyglossia}
  - \setmainlanguage{english}
  - \setotherlanguage{sanskrit}
  - \newfontfamily\englishfont[Ligatures=TeX]{Noto Sans}
  - \newfontfamily\sanskritfont[Script=Gujarati]{Noto Sans Gujarati}
---

# Introduction

This is a bilingual document that combines English and Gujarati. English will be the primary language, while Gujarati will be used for specific phrases or sentences.

# Gujarati Proverb

Here's a famous Gujarati proverb:

\begin{sanskrit}
સારા વિચારો જ સારા સંસ્કાર ઘડે છે.
\end{sanskrit}

Translation: Good thoughts shape good character.

# Gujarati Poem

Let's include a short Gujarati poem:

\begin{sanskrit}
ગુજરાતી ભાષા મધુર છે,\\
સૌને પ્રિય લાગે છે,\\
ગુજરાતી માંહી ગુંજે છે,\\
દેશ-વિદેશમાં પૂજે છે।
\end{sanskrit}

Translation:
Gujarati language is sweet,
Loved by all,
Resonating within Gujarat,
Revered in the country and abroad.

# Conclusion

This document showcases the beautiful combination of English and Gujarati languages. It demonstrates how we can seamlessly integrate Gujarati phrases and poetry into an English document.
