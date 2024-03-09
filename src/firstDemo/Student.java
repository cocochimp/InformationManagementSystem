package firstDemo;

public class Student {
    //成员变量
    private String name;
    private String id;
    private String startYear;
    private String profession;
    private String klass;
    private String number;

    //构造方法
    public Student(){}

    public Student(String name, String id, String startYear, String profession, String klass, String number) {
        this.name = name;
        this.id = id;
        this.startYear = startYear;
        this.profession = profession;
        this.klass = klass;
        this.number = number;
    }

    //成员变量

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public String getStartYear() {
        return startYear;
    }

    public void setStartYear(String startYear) {
        this.startYear = startYear;
    }

    public String getProfession() {
        return profession;
    }

    public void setProfession(String profession) {
        this.profession = profession;
    }

    public String getKlass() {
        return klass;
    }

    public void setKlass(String klass) {
        this.klass = klass;
    }

    public String getNumber() {
        return number;
    }

    public void setNumber(String number) {
        this.number = number;
    }






}
