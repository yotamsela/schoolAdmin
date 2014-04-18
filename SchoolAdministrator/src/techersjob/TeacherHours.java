package techersjob;

public class TeacherHours {
	private double percentage;
	private double privateHours;
	private double supportHours;
	
	
	
	public TeacherHours(double percentage, double privateHours,
			double supportHours) {
		super();
		this.percentage = percentage;
		this.privateHours = privateHours;
		this.supportHours = supportHours;
	}
	
	/**
	 * return a zero initialize TeacherAttribute instance, 
	 * this is done for teachers that does not participate in oz tmura.
	 * 
	 * @return
	 */
	public static TeacherHours getZeroHours(){
		return new TeacherHours(0, 0, 0);
	}


	public double getPercentage() {
		return percentage;
	}


	public double getPrivateHours() {
		return privateHours;
	}


	public double getSupportHours() {
		return supportHours;
	}
	
	
	
	

}
