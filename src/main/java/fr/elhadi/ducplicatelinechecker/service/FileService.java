package fr.elhadi.ducplicatelinechecker.service;

import com.fasterxml.jackson.databind.ObjectMapper;
import fr.elhadi.ducplicatelinechecker.model.Element;
import fr.elhadi.ducplicatelinechecker.model.KPILJ;
import fr.elhadi.ducplicatelinechecker.model.Root;
import fr.elhadi.ducplicatelinechecker.model.WorkingMode;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;
import java.util.function.Function;
import java.util.stream.Collectors;
import java.util.stream.Stream;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

public class FileService {
    
    
    private static final Logger logger = LogManager.getLogger();

    static Function<String, Element> consumeTxt = (line -> {
        String[] temp = line.split("\\+");
        Element element = new Element();
        if(temp.length > 0) {
            element.setDppType(temp[0]);
            element.setDppThird(temp[1]);
            element.setDppSubThird(temp[2]);
            element.setCustomerType(temp[3]);
            element.setCustomerThird(temp[4]);
            element.setCustomerSubTihrd(temp[5]);
            element.setFinalType(temp[6]);
            element.setFinalThird(temp[7]);
            element.setFinalSubThird(temp[8]);
            element.setItemNumber(temp[9]);
            element.setCurrentWeek(temp[10]);
            element.setCurrentYear(temp[11]);
            element.setReplenishment(temp[12]);
            element.setImplementation(temp[13]);
        }
        return element;});

    static Function<KPILJ, Element> consumeJson = (line -> {
        Element element = new Element();
            element.setDppType(line.getZTYPTIEDP());
            element.setDppThird(line.getZTIEDPP());
            element.setDppSubThird(line.getZSOUTIEDP());
            element.setCustomerType(line.getZTYPTIERS());
            element.setCustomerThird(line.getZTIERS());
            element.setCustomerSubTihrd(line.getZSOUTIERS());
            element.setFinalType(line.getZTYPTIEFCU());
            element.setFinalThird(line.getZTIEFCU());
            element.setFinalSubThird(line.getZSOUTIEFC());
            element.setItemNumber(line.getMATERIAL());
            element.setCurrentWeek(line.getFISCPER3());
            element.setCurrentYear(line.getCALYEAR());
            element.setReplenishment(line.getZDELAIREP());
            element.setImplementation(line.getZDELAIIMP());
        return element;});

    private FileService() {}
    
    public static boolean isFilePresent(String pathOfFile) {
        Path path = Paths.get(pathOfFile);
        File directory = path.toFile();

        return !directory.exists() || Objects.requireNonNull(directory.listFiles()).length <= 0;
    }
    public static List<Element> extractElementsFromDocument(WorkingMode workingMode, String pathOfFile) {
        List<Element> elementList = new ArrayList<>();
        
        try(Stream<Path> directory = Files.walk(Paths.get(pathOfFile))) {
            if(WorkingMode.OPERATION.equals(workingMode)) {
                elementList = parseOperationalFiles(directory);
            } else {
                elementList = parseTestingFiles(directory);
            }
            
        } catch (IOException e) {
            logger.debug("Erreur durant la récupération des fichiers", e);
        }
        return elementList;
    }
    
    private static List<Element> parseOperationalFiles(Stream<Path> directory) {
        List<KPILJ> linesOfFile = new ArrayList<>();
            directory.filter(Files::isRegularFile).forEach(file -> {
                ObjectMapper mapper = new ObjectMapper();
                Root root = null;
                try {
                    root = mapper.readValue(file.toFile(), Root.class);
                } catch (IOException e) {
                    logger.debug(String.format("Erreur durant le parsing du fichier %s", file.getFileName()), e);
                }
                if(root != null && CollectionUtils.isNotEmpty(root.getKPILJ())) {
                    linesOfFile.addAll(root.getKPILJ());
                }
            });
        return linesOfFile.stream().map(consumeJson).sorted().collect(Collectors.toList());
    }

    private static List<Element> parseTestingFiles(Stream<Path> directory) {
        List<String> linesOfFile = new ArrayList<>();
        directory.filter(Files::isRegularFile).forEach(file -> {
            try {
                linesOfFile.addAll(Files.readAllLines(file));
            } catch (IOException e) {
                logger.debug(String.format("Erreur durant le parsing du fichier %s", file.getFileName()), e);
            }
        });
        return linesOfFile.stream().map(consumeTxt).sorted().collect(Collectors.toList());
    }

    public static String lineIsPresentInDocument(String line, List<String> document) {
        if(document.contains(line)) {
            return null;
        } else {
            return line;
        }
    }
}
