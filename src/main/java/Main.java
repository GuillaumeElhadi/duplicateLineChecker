import fr.elhadi.ducplicatelinechecker.service.Element;
import fr.elhadi.ducplicatelinechecker.service.ExcelGenerator;
import fr.elhadi.ducplicatelinechecker.service.FileService;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.io.IOException;
import java.util.List;
import java.util.Scanner;
import java.util.stream.Collectors;

public class Main {

    public static void main(String... args) throws IOException {

        Logger logger = LogManager.getLogger(Main.class);

        Scanner sc = new Scanner(System.in);
        logger.info("Rentrez le chemin absolu de votre repertoire non optimisé");
        String regularDirectory = sc.nextLine();

        logger.info("Rentrez le chemin absolu de votre repertoire optimisé");
        String optimisedDirectory = sc.nextLine();

        if (FileService.isFilePresent(regularDirectory) || FileService.isFilePresent(optimisedDirectory)) {
            logger.error("fichiers non presents");
        } else {
            List<Element> elementsA = FileService.extractElementsFromDocument(regularDirectory);
            List<Element> elementsB = FileService.extractElementsFromDocument(optimisedDirectory);
            if (elementsA.equals(elementsB)) {
                logger.info("les deux fichiers sont identiques");
            } else if (elementsA.size() != elementsB.size()) {
                logger.info("Les fichiers ne contiennent pas le même nombre d'item");
            } else {
                logger.info("Il y a le même nombre d'articles mais des differences.");
                logger.info("Recherche des elements différents entre les deux flux");

                List<Element> elementsANotPresentInB = elementsA.parallelStream().filter(a -> !elementsB.contains(a)).collect(Collectors.toList());
                List<Element> elementsBNotPresentInA = elementsB.parallelStream().filter(b -> !elementsA.contains(b)).collect(Collectors.toList());

                ExcelGenerator.generateExcelFile(elementsANotPresentInB, elementsBNotPresentInA);

                logger.info("difference result");

            }
        }
    }

}
