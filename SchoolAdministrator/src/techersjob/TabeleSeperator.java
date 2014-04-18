package techersjob;

public enum TabeleSeperator{
	
	
	OZ_ABOVE_55(120,55,TableType.OZ,false),
	OZ_BETWEEN_50_54(54,50,TableType.OZ,false),
	OZ_BELOW_50(49,0,TableType.OZ,false),
	OZ_BELOW_50_Mothers(49,0,TableType.OZ,true),
	OZ_ABOVE_50_Mothers(120,50,TableType.OZ,true),
	
	OFEK_ABOVE_55(120,55,TableType.OFEK,false),
	OFEK_BETWEEN_50_54(54,50,TableType.OFEK,false),
	OFEK_BELOW_50(49,0,TableType.OFEK,false),
	OFEK_BELOW_50_Mothers(49,0,TableType.OFEK,true),
	OFEK_ABOVE_50_Mothers(120,50,TableType.OFEK,true);
	
	
	private int upperBound;
	private int lowerBound;
	private TableType type;
	private boolean isMother;
	
	private TabeleSeperator(int upperBound, int lowerBound,TableType type,boolean isMother) {
		this.upperBound = upperBound;
		this.lowerBound = lowerBound;
		this.type = type;
		this.isMother = isMother;
	}
	
	private boolean isInRange(int age){
		return (age <= upperBound && age >= lowerBound);
	}
	
	public static TabeleSeperator getRange(int age,TableType type,boolean isMother){
		for (TabeleSeperator var : TabeleSeperator.values()) {
			if(isMother == var.isMother && var.type == type && var.isInRange(age)){
				return var;
			}
        }	
		return null;
	}
	
	public enum TableType{OFEK,OZ};
}