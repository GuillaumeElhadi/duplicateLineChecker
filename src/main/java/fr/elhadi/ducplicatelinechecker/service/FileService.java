package fr.elhadi.ducplicatelinechecker.service;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;
import java.util.function.Function;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public class FileService {

    static Function<String, Element> consume = (line -> {
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
            element.setProvisions(temp[13]);
        }
        return element;});


    public static boolean isFilePresent(String pathOfFile) {
        Path path = Paths.get(pathOfFile);
        File directory = path.toFile();

        return !directory.exists() || Objects.requireNonNull(directory.listFiles()).length <= 0;
    }
    public static List<Element> extractElementsFromDocument(String pathOfFile) {
        List<String> linesOfFile = new ArrayList<>();
        try(Stream<Path> directory = Files.walk(Paths.get(pathOfFile))) {
            directory.filter(Files::isRegularFile).forEach(file -> {
                try {
                    linesOfFile.addAll(Files.readAllLines(file));
                } catch (IOException ioException) {
                    ioException.printStackTrace();
                }
            });
        } catch (IOException e) {
            e.printStackTrace();
        }
        return linesOfFile.stream().map(consume).sorted().collect(Collectors.toList());
    }

    public static String lineIsPresentInDocument(String line, List<String> document) {
        if(document.contains(line)) {
            return null;
        } else {
            return line;
        }
    }

    public static String itemIsPresentButDifferentInDocument(String line, List<String> document) {
        String [] element = line.split("/+");
        return null;
    }

}
