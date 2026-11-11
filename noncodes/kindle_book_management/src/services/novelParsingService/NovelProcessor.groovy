package services.novelParsingService


import services.kndiyLibraries.DataStructure
import org.apache.poi.xwpf.usermodel.XWPFParagraph

class NovelProcessor {
    List<String> allParagraphs

    boolean isSkippable
    boolean hasSkipSigns
    boolean chapterTitleFound
    String skippableStartSign
    String skippableEndSign
    String bookName
    String authorName
    String chapterStartsWithText
    Set skippableTexts = [
            "Prev Chapter",
            "Next Chapter",
            "Reset",
            "Generate Font Colors",
            "Reset",
            ""
    ]

    LinkedList possibleChapterSeparators = [
            " ", ":", "-"
    ]

    int totalChapterCount
    int currentChapterNumber
    String currentChapterTracker

    String outputPath
    String currentParagraphText
    String lastParagraphText
    String templatePath

    XWPFParagraph currentParagraph

    TreeMap<String, Map> paragraphsByTypeByChapterNumber

    NovelProcessor(String sourcePath, String outputPath, String templatePath,
                   String skippableStartSign, String skippableEndSign,
                   String chapterStartsWithText,
                   String bookName, String authorName,
                   Set skippableTexts = null) {
        File source = new File(sourcePath)

        allParagraphs = [ ]
        (source.isDirectory() ? source.listFiles().toList() : [ source ]).each { file ->
            List<String> paragraphs = getDocumentParagraphs(file)
            if (paragraphs) {
                allParagraphs.addAll(paragraphs)
            }
        }

        this.skippableStartSign = skippableStartSign
        this.skippableEndSign = skippableEndSign
        this.chapterStartsWithText = chapterStartsWithText
        this.bookName = bookName
        this.authorName = authorName
        this.outputPath = outputPath
        this.templatePath = templatePath
        // without skip signs there is no end-sign marker between chapters, so a new chapter number starts one
        hasSkipSigns = skippableStartSign && skippableEndSign
        isSkippable = hasSkipSigns
        chapterTitleFound = false
        totalChapterCount = 0
        for (String skippableText : skippableTexts) {
            this.skippableTexts.add(skippableText)
        }
        String spaces = ""
        for (int i = 0; i < 100; i ++) {
            spaces += " "
            this.skippableTexts.add(spaces)
        }
        paragraphsByTypeByChapterNumber = new TreeMap<>()
        updateParagraphsByTypeByChapterNumber()

        new NovelWriteToDocx(this).printOutputDoc()
    }

    List<String> getDocumentParagraphs(File file) {
        String fileType = file.getName().contains(".") ? file.getName().substring(file.getName().lastIndexOf(".") + 1) : null

        if (fileType == "docx") {
            return new NovelParserXwpf(file.getAbsolutePath()).getParagraphTexts()
        } else if (fileType == "epub") {
            return new NovelParserEpub(file.getAbsolutePath()).getParagraphTexts()
        }

        return null
    }

    void updateParagraphsByTypeByChapterNumber() {
        for (String paragraph : allParagraphs) {
            currentParagraphText = paragraph.trim()
            updateIsSkippable()
            if (isSkippable || skippableTexts.contains(currentParagraphText)) {
                continue
            }

            int previousChapterNumber = currentChapterNumber
            boolean isChapterParagraph = checkIfHasChapterNumber()
            if (!hasSkipSigns && isChapterParagraph && currentChapterNumber != previousChapterNumber) {
                chapterTitleFound = false
            }

            currentChapterTracker = String.format("%04d", currentChapterNumber)
            if (chapterTitleFound && isChapterParagraph) {
                continue
            }
            if (!chapterTitleFound && isChapterParagraph) {
                println("Parsing Chapter: ${currentChapterTracker}")
                Map paragraphsByType = DataStructure.getOrCreateObject(
                        paragraphsByTypeByChapterNumber, currentChapterTracker, DataStructure.DATA_STRUCTURE_MAP
                )
                String titleString = "${currentParagraphText - "Novel "}"?.trim()
                paragraphsByType["title"] = [ titleString, authorName ]
                lastParagraphText = authorName
                chapterTitleFound = true
            }
            else if (currentParagraphText != bookName && !isChapterParagraph) {
                Map paragraphsByType = DataStructure.getOrCreateObject(
                        paragraphsByTypeByChapterNumber, currentChapterTracker, DataStructure.DATA_STRUCTURE_MAP
                )
                List containedParagraphs = DataStructure.getOrCreateObject(
                        paragraphsByType, "body", DataStructure.DATA_STRUCTURE_LIST
                )
                containedParagraphs.add(currentParagraphText)
                lastParagraphText = currentParagraphText
            }
        }
    }

    void updateIsSkippable() {
        if (!hasSkipSigns) {
            return
        }
        if (currentParagraphText.startsWith(skippableStartSign)) {
            isSkippable = true
        }
        if (currentParagraphText.startsWith(skippableEndSign)) {
            isSkippable = false
            chapterTitleFound = false
            totalChapterCount ++
        }
    }

    boolean checkIfHasChapterNumber() {
        if (currentParagraphText.startsWith(chapterStartsWithText)) {
            Integer chapterNumber = findChapterNumber()
            if (chapterNumber) {
                currentChapterNumber = chapterNumber
                return true
            }
        }

        return false
    }

    Integer findChapterNumber() {
        LinkedList allCompos = [ ]
        for (String separator : possibleChapterSeparators) {
            findAllToCheckTexts(allCompos, separator, currentParagraphText)
        }

        for (String text : allCompos) {
            try {
                return text.trim()
                        .toBigInteger()
                        .toInteger()
            }
            catch (ignored) {}
        }

        return null
    }

    LinkedList findAllToCheckTexts(LinkedList allCompos, String separator, String textToSplit) {
        if (!allCompos) {
            allCompos.addAll(textToSplit.split(separator))
        }
        else {
            for (int idx = 0; idx < allCompos.size(); idx ++) {
                String text = allCompos.remove(idx)
                List compos = text.split(separator)
                allCompos.addAll(idx, compos)
            }
        }

        return allCompos
    }
}
