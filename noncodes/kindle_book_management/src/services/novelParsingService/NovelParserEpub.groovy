package services.novelParsingService

import groovy.xml.XmlSlurper

import java.util.regex.Pattern
import java.util.zip.ZipFile

/**
 * Reads an epub in spine order and flattens it into paragraph texts: for every page that carries a
 * chapter title (h2.chapter-title) the title comes first, then the page's paragraphs.
 * Cover, information and table-of-contents pages have no chapter title and are left out.
 */
class NovelParserEpub extends NovelParser {
    private static final Pattern CHAPTER_TITLE = Pattern.compile('chapter-title">(.*?)</h2>', Pattern.DOTALL)
    private static final Pattern PARAGRAPH = Pattern.compile('<p[^>]*>(.*?)</p>', Pattern.DOTALL)
    private static final Pattern TAG = Pattern.compile('<[^>]+>')
    private static final Pattern CONTROL_CHARACTER = Pattern.compile('[\\x00-\\x08\\x0b\\x0c\\x0e-\\x1f]')

    NovelParserEpub(String fileAbsolutePath) {
        super(fileAbsolutePath)
    }

    @Override
    List<String> getParagraphTexts() {
        List<String> paragraphs = [ ]
        ZipFile epub = new ZipFile(fileAbsolutePath)
        try {
            getSpinePagePaths(epub).each { String pagePath ->
                String page = readEntry(epub, pagePath)
                def titleMatcher = CHAPTER_TITLE.matcher(page)
                if (!titleMatcher.find()) {
                    return
                }

                paragraphs.add(toText(titleMatcher.group(1)))
                def paragraphMatcher = PARAGRAPH.matcher(page)
                while (paragraphMatcher.find()) {
                    paragraphs.add(toText(paragraphMatcher.group(1)))
                }
            }
        }
        finally {
            epub.close()
        }

        return paragraphs
    }

    private static List<String> getSpinePagePaths(ZipFile epub) {
        def container = new XmlSlurper().parseText(readEntry(epub, "META-INF/container.xml"))
        String opfPath = container.rootfiles.rootfile.@"full-path".text()
        def opf = new XmlSlurper().parseText(readEntry(epub, opfPath))

        Map<String, String> hrefById = [ : ]
        opf.manifest.item.each { hrefById[it.@id.text()] = it.@href.text() }

        String baseFolder = opfPath.contains("/") ? opfPath.substring(0, opfPath.lastIndexOf("/") + 1) : ""
        return opf.spine.itemref.collect { baseFolder + hrefById[it.@idref.text()] }
    }

    private static String readEntry(ZipFile epub, String path) {
        return epub.getInputStream(epub.getEntry(path)).getText("UTF-8")
    }

    private static String toText(String htmlFragment) {
        String text = TAG.matcher(htmlFragment).replaceAll("")
        text = text.replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", "\"")
                .replace("&apos;", "'").replace("&nbsp;", " ").replace("&amp;", "&")
        return CONTROL_CHARACTER.matcher(text).replaceAll("").trim()
    }
}
