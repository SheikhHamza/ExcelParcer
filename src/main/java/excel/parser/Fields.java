package excel.parser;

public class Fields{
    private String refinedName;
    private String originalName;

    public Fields(String refinedName, String originalName) {
        this.refinedName = refinedName;
        this.originalName = originalName;
    }

    public String getRefinedName() {
        return refinedName;
    }

    public void setRefinedName(String refinedName) {
        this.refinedName = refinedName;
    }

    public String getOriginalName() {
        return originalName;
    }

    public void setOriginalName(String originalName) {
        this.originalName = originalName;
    }
}