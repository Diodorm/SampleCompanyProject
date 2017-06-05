public class Customer {

    private String name;
    private String maths;
    private String science;
    private String english;

    public Customer(){}

    public Customer(String name, String maths, String science, String english) {
        this.name = name;
        this.maths = maths;
        this.science = science;
        this.english = english;
    }

    //getters and setter..

    @Override
    public String toString() {
        return name+ ": Maths "+maths+ " Science "+science+" English "+english;
    }
}
