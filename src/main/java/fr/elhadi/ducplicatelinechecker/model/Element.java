package fr.elhadi.ducplicatelinechecker.model;

import java.util.Objects;

public class Element implements Comparable<Element>{

    private String dppType;
    private String dppThird;
    private String dppSubThird;

    private String customerType;
    private String customerThird;
    private String customerSubTihrd;

    private String finalType;
    private String finalThird;
    private String finalSubThird;

    private String itemNumber;

    private String currentWeek;
    private String currentYear;


    private String replenishment;
    private String implementation;

    public String getDppType() {
        return dppType;
    }

    public void setDppType(String dppType) {
        this.dppType = dppType;
    }

    public String getDppThird() {
        return dppThird;
    }

    public void setDppThird(String dppThird) {
        this.dppThird = dppThird;
    }

    public String getDppSubThird() {
        return dppSubThird;
    }

    public void setDppSubThird(String dppSubThird) {
        this.dppSubThird = dppSubThird;
    }

    public String getCustomerType() {
        return customerType;
    }

    public void setCustomerType(String customerType) {
        this.customerType = customerType;
    }

    public String getCustomerThird() {
        return customerThird;
    }

    public void setCustomerThird(String customerThird) {
        this.customerThird = customerThird;
    }

    public String getCustomerSubTihrd() {
        return customerSubTihrd;
    }

    public void setCustomerSubTihrd(String customerSubTihrd) {
        this.customerSubTihrd = customerSubTihrd;
    }

    public String getFinalType() {
        return finalType;
    }

    public void setFinalType(String finalType) {
        this.finalType = finalType;
    }

    public String getFinalThird() {
        return finalThird;
    }

    public void setFinalThird(String finalThird) {
        this.finalThird = finalThird;
    }

    public String getFinalSubThird() {
        return finalSubThird;
    }

    public void setFinalSubThird(String finalSubThird) {
        this.finalSubThird = finalSubThird;
    }

    public String getItemNumber() {
        return itemNumber;
    }

    public void setItemNumber(String itemNumber) {
        this.itemNumber = itemNumber;
    }

    public String getCurrentWeek() {
        return currentWeek;
    }

    public void setCurrentWeek(String currentWeek) {
        this.currentWeek = currentWeek;
    }

    public String getCurrentYear() {
        return currentYear;
    }

    public void setCurrentYear(String currentYear) {
        this.currentYear = currentYear;
    }

    public String getReplenishment() {
        return replenishment;
    }

    public void setReplenishment(String replenishment) {
        this.replenishment = replenishment;
    }

    public String getImplementation() {
        return implementation;
    }

    public void setImplementation(String implementation) {
        this.implementation = implementation;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        Element element = (Element) o;
        return Objects.equals(getDppType(), element.getDppType()) &&
                Objects.equals(getDppThird(), element.getDppThird()) &&
                Objects.equals(getDppSubThird(), element.getDppSubThird()) &&
                Objects.equals(getCustomerType(), element.getCustomerType()) &&
                Objects.equals(getCustomerThird(), element.getCustomerThird()) &&
                Objects.equals(getCustomerSubTihrd(), element.getCustomerSubTihrd()) &&
                Objects.equals(getFinalType(), element.getFinalType()) &&
                Objects.equals(getFinalThird(), element.getFinalThird()) &&
                Objects.equals(getFinalSubThird(), element.getFinalSubThird()) &&
                Objects.equals(getItemNumber(), element.getItemNumber()) &&
                Objects.equals(getCurrentWeek(), element.getCurrentWeek()) &&
                Objects.equals(getCurrentYear(), element.getCurrentYear()) &&
                Objects.equals(getReplenishment(), element.getReplenishment()) &&
                Objects.equals(getImplementation(), element.getImplementation());
    }

    @Override
    public int hashCode() {
        return Objects.hash(getDppType(), getDppThird(), getDppSubThird(), getCustomerType(), getCustomerThird(), getCustomerSubTihrd(), getFinalType(), getFinalThird(), getFinalSubThird(), getItemNumber(), /*getCurrentWeek(),*/ getCurrentYear(), getReplenishment(), getImplementation());
    }

    @Override
    public int compareTo(Element o) {
        if(this.getItemNumber().equals(o.getItemNumber())) {
            return 0;
        } else if(Integer.parseInt(this.getItemNumber()) < Integer.parseInt(o.getItemNumber())) {
            return -1;
        } else {
            return 1;
        }
    }
}
