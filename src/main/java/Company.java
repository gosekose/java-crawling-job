public class Company {
    private String business;
    private String classification;

    public Company(String business, String classification) {
        this.business = business;
        this.classification = classification;
    }

    public String getBusiness() {
        return business;
    }

    public void setBusiness(String business) {
        this.business = business;
    }

    public String getClassification() {
        return classification;
    }

    public void setClassification(String classification) {
        this.classification = classification;
    }
}
