package bagruyot;
import org.eclipse.swt.SWT;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;
import org.eclipse.swt.widgets.Label;
import org.eclipse.swt.layout.GridLayout;
import org.eclipse.swt.widgets.Text;
import org.eclipse.swt.layout.GridData;

public class PasswordShell extends Shell {
	private Text text;

	/**
	 * Launch the application.
	 * @param args
	 */
	public static void main(String args[]) {
		try {
			Display display = Display.getDefault();
			PasswordShell shell = new PasswordShell(display);
			shell.open();
			shell.layout();
			while (!shell.isDisposed()) {
				if (!display.readAndDispatch()) {
					display.sleep();
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * Create the shell.
	 * @param display
	 */
	public PasswordShell(Display display) {
		super(display, SWT.SHELL_TRIM | SWT.RIGHT_TO_LEFT);
		setLayout(new GridLayout(1, false));
		
		Label label = new Label(this, SWT.NONE);
		label.setLayoutData(new GridData(SWT.CENTER, SWT.CENTER, false, false, 1, 1));
		label.setFont(SWTResourceManager.getFont("Tahoma", 15, SWT.NORMAL));
		label.setText("\u05D1\u05E8\u05D5\u05DB\u05D9\u05DD \u05D4\u05D1\u05D0\u05D9\u05DD \u05DC\u05EA\u05D5\u05DB\u05E0\u05EA");
		
		Label lblSchoolAdministrator = new Label(this, SWT.NONE);
		lblSchoolAdministrator.setFont(SWTResourceManager.getFont("Tahoma", 22, SWT.BOLD));
		lblSchoolAdministrator.setLayoutData(new GridData(SWT.CENTER, SWT.CENTER, false, false, 1, 1));
		lblSchoolAdministrator.setText("School Administrator");
		
		Label lblNewLabel = new Label(this, SWT.NONE);
		lblNewLabel.setFont(SWTResourceManager.getFont("Tahoma", 12, SWT.BOLD));
		lblNewLabel.setLayoutData(new GridData(SWT.CENTER, SWT.CENTER, false, false, 1, 1));
		lblNewLabel.setText("\u05DC\u05EA\u05D7\u05D9\u05DC\u05EA \u05E4\u05E2\u05D9\u05DC\u05D5\u05EA \u05D4\u05D6\u05D9\u05E0\u05D5 \u05E1\u05D9\u05E1\u05DE\u05D4:");
		
		text = new Text(this, SWT.BORDER);
		text.setText("                    ");
		text.setLayoutData(new GridData(SWT.CENTER, SWT.CENTER, true, false, 1, 1));
		createContents();
	}

	/**
	 * Create contents of the shell.
	 */
	protected void createContents() {
		setText("\u05D4\u05D6\u05E0\u05EA \u05E1\u05D9\u05E1\u05DE\u05D4");
		setSize(450, 161);

	}

	@Override
	protected void checkSubclass() {
		// Disable the check that prevents subclassing of SWT components
	}
}
