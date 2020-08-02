package enumeration;

public enum GenderEnum {
    MALE("男","1"),
    FEMALE("女", "2")
    ;
    private String info;
    private String value;

    GenderEnum(String info, String value){
        this.info = info;
        this.value = value;
    }

    public String getInfo(){
        return info;
    }

}
