package fr.elhadi.ducplicatelinechecker.model;


import com.fasterxml.jackson.annotation.JsonProperty;
import java.util.List;

public class Root{
    @JsonProperty("KPILJ")
    public List<KPILJ> getKPILJ() {
        return this.kPILJ; }
    public void setKPILJ(List<KPILJ> kPILJ) {
        this.kPILJ = kPILJ; }
    List<KPILJ> kPILJ;
}
