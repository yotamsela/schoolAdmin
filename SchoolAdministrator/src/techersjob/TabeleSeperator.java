package techersjob;

public enum TabeleSeperator{
	
	
	OZ_ABOVE_55(120,55,TableType.OZ),
	OZ_BETWEEN_50_54(54,50,TableType.OZ),
	OZ_BELOW_50(49,0,TableType.OZ),
	
	OFEK_ABOVE_55(120,55,TableType.OFEK),
	OFEK_BETWEEN_50_54(54,50,TableType.OFEK),
	OFEK_BELOW_50(49,0,TableType.OFEK);
	
	
	private int upperBound;
	private int lowerBound;
	private TableType type;
	
	private TabeleSeperator(int upperBound, int lowerBound,TableType type) {
		this.upperBound = upperBound;
		this.lowerBound = lowerBound;
		this.type = type;
	}
	
	private boolean isInRange(int age){
		return (age <= upperBound && age >= lowerBound);
	}
	
	public static TabeleSeperator getRange(int age,TableType type){
		for (TabeleSeperator var : TabeleSeperator.values()) {
			if(var.type == type && var.isInRange(age)){
				return var;
			}
        }	
		return null;
	}
	
	public enum TableType{OFEK,OZ};
}