import java.util.LinkedList;

public class Rubric {
	
	public void addCategory(double weight, String name, int numAssignments) {
		Category c = new Category(name, weight, numAssignments);
		_categories.add(c);
	}
	
	public Category getCategory(int position) {
		return _categories.get(position);
	}
	
	public LinkedList<Category> getCategories() {
		return _categories;
	}
	
	public LinkedList<Category> _categories = new LinkedList<Category>();
}