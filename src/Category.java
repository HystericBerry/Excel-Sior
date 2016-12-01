
public class Category {
	public Category(String name, double weight, int numAssignments) {
		_weight = weight;
		_name = name;
		_numAgmt = numAssignments;
	}
	
	public String getName() {
		return _name;
	}
	
	public double getWeight() {
		return _weight;
	}
	
	public int getNumAssign() {
		return _numAgmt;
	}
	
	String _name; // name of category (e.g. Homework)
	double _weight; // weight of grade (e.g. 25.0%)
	int _numAgmt; // number of assignments (e.g. 10 homeworks)
}