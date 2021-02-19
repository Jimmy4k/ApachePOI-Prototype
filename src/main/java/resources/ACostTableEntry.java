/**
 *
 */
package resources;

public class ACostTableEntry {


private String propertyType = "type";
private double propertyCost = 0.00;
private double PROPERTY_COSTWITHMARGIN = 0.00;
private double PROPERTY_COSTMARGIN = 0.00;
private double PROPERTY_MARGIN = 0.00;

	public String getType() {
		return propertyType;
	}

	public void setType(String pROPERTY_TYPE) {
		propertyType = pROPERTY_TYPE;
	}

	public double getCost() {
		return propertyCost;
	}

	public void setCost(double pROPERTY_COST) {
		propertyCost = pROPERTY_COST;
	}

	public double getCostWithMargin() {
		return PROPERTY_COSTWITHMARGIN;
	}

	public void setCostWithMargin(double pROPERTY_COSTWITHMARGIN) {
		PROPERTY_COSTWITHMARGIN = pROPERTY_COSTWITHMARGIN;
	}

	public double getCostMargin() {
		return PROPERTY_COSTMARGIN;
	}

	public void setCostMargin(double pROPERTY_COSTMARGIN) {
		PROPERTY_COSTMARGIN = pROPERTY_COSTMARGIN;
	}

	public double getMargin() {
		return PROPERTY_MARGIN;
	}

	public void setMargin(double pROPERTY_MARGIN) {
		PROPERTY_MARGIN = pROPERTY_MARGIN;
	}

}
