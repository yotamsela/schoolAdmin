package bagruyot;
import java.io.File;
import org.eclipse.swt.SWT;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.FileDialog;
import org.eclipse.swt.widgets.Shell;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;
import org.eclipse.swt.layout.GridLayout;
import org.eclipse.swt.widgets.Label;
import org.eclipse.swt.layout.GridData;


public class MainShell extends Shell {

	private Shell me;
	private String filePath;


	/**
	 * Create the shell.
	 * @param display
	 */
	public MainShell(Display display) {
		super(display, SWT.SHELL_TRIM | SWT.RIGHT_TO_LEFT);
		me = this;
		filePath = "";

		setSize(495, 299);
		setLayout(new GridLayout(6, false));
		new Label(this, SWT.NONE);
		
		Label lblNewLabel = new Label(this, SWT.WRAP | SWT.CENTER);
		lblNewLabel.setLayoutData(new GridData(SWT.CENTER, SWT.CENTER, false, false, 5, 2));
		lblNewLabel.setFont(SWTResourceManager.getFont("Tahoma", 22, SWT.BOLD));
		lblNewLabel.setAlignment(SWT.CENTER);
		lblNewLabel.setText("School Administrator");
		new Label(this, SWT.NONE);
		new Label(this, SWT.NONE);
		
		Label lblNewLabel_1 = new Label(this, SWT.CENTER);
		lblNewLabel_1.setLayoutData(new GridData(SWT.CENTER, SWT.CENTER, false, false, 5, 1));
		lblNewLabel_1.setAlignment(SWT.CENTER);
		lblNewLabel_1.setFont(SWTResourceManager.getFont("Tahoma", 14, SWT.BOLD));
		lblNewLabel_1.setText("\u05D2\u05E8\u05E1\u05D4 \u05E0\u05D9\u05E1\u05D9\u05D5\u05E0\u05D9\u05EA");
		new Label(this, SWT.NONE);
		new Label(this, SWT.NONE);
		new Label(this, SWT.NONE);
		new Label(this, SWT.NONE);
		new Label(this, SWT.NONE);
		new Label(this, SWT.NONE);
		new Label(this, SWT.NONE);
		new Label(this, SWT.NONE);
		new Label(this, SWT.NONE);
		new Label(this, SWT.NONE);
		new Label(this, SWT.NONE);
		new Label(this, SWT.NONE);
		new Label(this, SWT.NONE);
		
		Label filesLabel = new Label(this, SWT.RIGHT);
		filesLabel.setAlignment(SWT.LEFT);
		filesLabel.setText("\u05E7\u05D1\u05E6\u05D9\u05DD:");
		new Label(this, SWT.NONE);
		
		Label classesLabel = new Label(this, SWT.NONE);
		classesLabel.setText("\u05DB\u05D9\u05EA\u05D5\u05EA:             ");
		
		Label fileLabel = new Label(this, SWT.NONE);
		fileLabel.setText("\u05E7\u05D5\u05D1\u05E5:");
		
		Button button = new Button(this, SWT.CENTER);
		button.setLayoutData(new GridData(SWT.CENTER, SWT.CENTER, false, false, 1, 1));
		button.setText("\u05E9\u05E0\u05D4 \u05E1\u05D9\u05E1\u05DE\u05D4");
		new Label(this, SWT.NONE);
		
		final Button teachersButton = new Button(this, SWT.CHECK);
		teachersButton.setSelection(true);
		teachersButton.setText("\u05DE\u05DE\u05D5\u05E6\u05E2\u05D9 \u05DE\u05D5\u05E8\u05D9\u05DD");
		new Label(this, SWT.NONE);
		
		final Button grade10Button = new Button(this, SWT.CHECK);
		grade10Button.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent arg0) {
			}
		});
		grade10Button.setText("\u05D9'");
		
		final Label fileNameLabel = new Label(this, SWT.NONE);
		fileNameLabel.setText("\u05DC\u05D0 \u05E0\u05D1\u05D7\u05E8");
		
		Button button_1 = new Button(this, SWT.NONE);
		button_1.setLayoutData(new GridData(SWT.CENTER, SWT.CENTER, false, false, 1, 1));
		button_1.setText("\u05E2\u05D3\u05DB\u05DF \u05E9\u05D0\u05DC\u05D5\u05E0\u05D9\u05DD \u05DC\u05DE\u05E7\u05E6\u05D5\u05E2");
		new Label(this, SWT.NONE);
		
		final Button studentsButton = new Button(this, SWT.CHECK | SWT.RIGHT);
		studentsButton.setAlignment(SWT.LEFT);
		studentsButton.setSelection(true);
		studentsButton.setText("\u05DE\u05DE\u05D5\u05E6\u05E2\u05D9 \u05EA\u05DC\u05DE\u05D9\u05D3\u05D9\u05DD");
		new Label(this, SWT.NONE);
		
		final Button grade11Button = new Button(this, SWT.CHECK);
		grade11Button.setText("\u05D9\"\u05D0");
		
		Button chooseFileButton = new Button(this, SWT.NONE);
		chooseFileButton.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent arg0) {
				FileDialog fd=new FileDialog(me ,SWT.OPEN);
				fd.setText("בחר קובץ");
				fd.setFilterPath(new File("").getAbsolutePath());
				//TODO fix filter
				String[] filterExt = {"*"};
				fd.setFilterExtensions(filterExt);
				String fileName = fd.open();
				if(fileName != null){
					filePath = fileName;
					fileNameLabel.setText(fileName.substring(fileName.lastIndexOf('\\') + 1, fileName.lastIndexOf('.')));
				}
			}
		});
		chooseFileButton.setText("\u05D1\u05D7\u05E8 \u05E7\u05D5\u05D1\u05E5");
		
		Button button_2 = new Button(this, SWT.NONE);
		button_2.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				ExamsReader.outputFile();
			}
		});
		button_2.setLayoutData(new GridData(SWT.CENTER, SWT.CENTER, false, false, 1, 1));
		button_2.setText("\u05E7\u05D1\u05DC \u05E8\u05E9\u05D9\u05DE\u05EA \u05E9\u05D0\u05DC\u05D5\u05E0\u05D9\u05DD");
		new Label(this, SWT.NONE);
		
		final Button errorsButton = new Button(this, SWT.CHECK);
		errorsButton.setText("\u05E9\u05D2\u05D9\u05D0\u05D5\u05EA");
		new Label(this, SWT.NONE);
		
		final Button grade12Button = new Button(this, SWT.CHECK);
		grade12Button.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent arg0) {
			}
		});
		grade12Button.setSelection(true);
		grade12Button.setText("\u05D9\"\u05D1");
		new Label(this, SWT.NONE);
		new Label(this, SWT.NONE);
		new Label(this, SWT.NONE);
		
		Button startButton = new Button(this, SWT.NONE);
		startButton.setLayoutData(new GridData(SWT.FILL, SWT.FILL, false, false, 4, 1));
		startButton.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent arg0) {
				if(filePath.equals("")){
					//TODO add error message
				}
				else{
					SchoolProgram.fullCalculation(grade10Button.getSelection(), grade11Button.getSelection(), grade12Button.getSelection(),
							teachersButton.getSelection(), studentsButton.getSelection(), errorsButton.getSelection(), filePath);
					me.close();
				}
			}
		});
		startButton.setText("\u05D4\u05EA\u05D7\u05DC \u05E0\u05D9\u05EA\u05D5\u05D7 \u05E0\u05EA\u05D5\u05E0\u05D9\u05DD");
		new Label(this, SWT.NONE);
		createContents();
	}

	/**
	 * Create contents of the shell.
	 */
	protected void createContents() {
		setText("SWT Application");

	}

	@Override
	protected void checkSubclass() {
		// Disable the check that prevents subclassing of SWT components
	}
}
