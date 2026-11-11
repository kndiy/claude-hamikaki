import services.novelParsingService.NovelProcessor

// args: epub, outputDocx, bookName, authorName [chapterPrefix=Chapter ] [templatePath] [skipText1|skipText2|...]
if (args.length < 4) {
    println("Usage: run_epub_to_docx.sh <in.epub> <out.docx> <bookName> <authorName> [chapterPrefix] [templatePath] [skipTexts separated by |]")
    System.exit(1)
}
String template = args.length > 5 && args[5] ? args[5] : new File(getClass().protectionDomain.codeSource.location.toURI()).parentFile.absolutePath + "/template/Template.docx"
Set skips = args.length > 6 && args[6] ? args[6].split("\\|") as Set : [ ] as Set
new NovelProcessor(args[0], args[1], template, null, null,
        args.length > 4 && args[4] ? args[4] : "Chapter ", args[2], args[3], skips)
