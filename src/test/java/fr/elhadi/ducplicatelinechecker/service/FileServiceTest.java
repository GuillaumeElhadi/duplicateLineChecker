package fr.elhadi.ducplicatelinechecker.service;

import fr.elhadi.ducplicatelinechecker.model.Element;
import fr.elhadi.ducplicatelinechecker.model.WorkingMode;
import junit.framework.TestCase;

import java.util.Arrays;
import java.util.List;

public class FileServiceTest extends TestCase {

    public void testExtractElementsFromDocument() {
        List<Element> linesOfThatFile = FileService.extractElementsFromDocument(WorkingMode.TEST, "src/test/resources/source.txt");
        assertNotNull(linesOfThatFile);
    }

    public void testLineIsPresentInDocument() {
        String line = "identical";
        List<String> document = Arrays.asList("i", "id", "ide", "iden", "ident", "identi", "identic", "identica", "identical");

        assertNull(FileService.lineIsPresentInDocument(line, document));
    }

    public void testLineIsNotPresentInDocument() {
        String line = "identicale";
        List<String> document = Arrays.asList("i", "id", "ide", "iden", "ident", "identi", "identic", "identica", "identical");

        assertEquals(line, FileService.lineIsPresentInDocument(line, document));
    }

}