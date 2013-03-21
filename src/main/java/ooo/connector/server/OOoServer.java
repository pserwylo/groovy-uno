package ooo.connector.server;

import com.sun.star.comp.helper.BootstrapException;
import com.sun.star.lib.util.NativeLibraryLoader;
import java.io.BufferedReader;
import java.io.File;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.IOException;
import java.io.PrintStream;
import java.net.MalformedURLException;
import java.net.URL;
import java.net.URLClassLoader;
import java.util.ArrayList;
import java.util.List;

/**
 * Starts and stops an OOo server.
 * 
 * Most of the source code in this class has been taken from the Java class
 * "Bootstrap.java" (Revision: 1.15) from the UDK projekt (Uno Software Develop-
 * ment Kit) from OpenOffice.org (http://udk.openoffice.org/). The source code
 * is available for example through a browser based online version control
 * access at http://udk.openoffice.org/source/browse/udk/. The Java class
 * "Bootstrap.java" is there available at
 * http://udk.openoffice.org/source/browse/udk/javaunohelper/com/sun/star/comp/helper/Bootstrap.java?view=markup
 */
public class OOoServer {

	public static final String ACCEPT_SOCKET = "-accept=socket,host=localhost,port=8100;urp;";

	public static final String ACCEPT_PIPE   = "-accept=pipe,name=oooPipe;urp;";

    /** The OOo server process. */
    private Process       oooProcess;

    /** Points to the relevant installations 'program' folder and the 'soffice' executable. */
    private OOoServerPath serverPath;

    /** The options for starting the OOo server. */
    private List<String>  oooOptions = new ArrayList<String>();

    /**
     * Constructs an OOo server which uses the folder of the OOo installation
     * containing the soffice executable and a list of default options to start
     * OOo.
     * 
     * @param serverPath Where to find the 'soffice' bin and the 'program' folder.
     */
    public OOoServer( OOoServerPath serverPath ) {

        this.oooProcess = null;
        this.serverPath = serverPath;
        this.oooOptions = getDefaultOOoOptions();
    }

    /**
     * Constructs an OOo server which uses the folder of the OOo installation
     * containing the soffice executable and a given list of options to start
     * OOo.
     * 
     * @param   serverPath   Where to find the 'soffice' bin and the 'program' folder.
     * @param   oooOptions   The list of options
     */
    public OOoServer(OOoServerPath serverPath, List<String> oooOptions) {

        this.oooProcess = null;
        this.serverPath = serverPath;
        this.oooOptions = oooOptions;
    }

    /**
     * Starts an OOo server which uses the specified accept option.
     * 
     * The accept option can be used for two different types of connections:
     * 1) The socket connection
     * 2) The named pipe connection
     * 
     * To create a socket connection a host and port must be provided.
     * For example using the host "localhost" and the port "8100" the
     * accept option looks like this:
     * - accept option    : -accept=socket,host=localhost,port=8100;urp;
     * 
     * To create a named pipe a pipe name must be provided. For example using
     * the pipe name "oooPipe" the accept option looks like this:
     * - accept option    : -accept=pipe,name=oooPipe;urp;
     * 
     * @param   oooAcceptOption      The accept option
     */
    public void start(String oooAcceptOption) throws BootstrapException, IOException {

        // find office executable relative to this class's class loader
        String sOffice = System.getProperty("os.name").startsWith("Windows")? "soffice.exe": "soffice";

        URL[] oooExecFolderURL = new URL[] {new File( serverPath.getProgramPath() ).toURI().toURL()};
        URLClassLoader loader = new URLClassLoader(oooExecFolderURL);
        File fOffice = NativeLibraryLoader.getResource(loader, sOffice);
        if (fOffice == null)
            throw new BootstrapException("Could not find 'soffice' executable at " + serverPath.getProgramPath() );

        // create call with arguments
        int arguments = oooOptions.size()+1;
        if (oooAcceptOption != null)
            arguments++;

        String[] oooCommand = new String[arguments];
        oooCommand[0] = fOffice.getPath();

        for (int i = 0; i < oooOptions.size(); i++) {
            oooCommand[i+1] = oooOptions.get(i);
        }

        if (oooAcceptOption != null)
            oooCommand[arguments-1] = oooAcceptOption;

        // start office process
        oooProcess = Runtime.getRuntime().exec(oooCommand);

        pipe(oooProcess.getInputStream(), System.out, "CO> ");
        pipe(oooProcess.getErrorStream(), System.err, "CE> ");
    }

    /**
     * Kills the OOo server process from the previous start.
     * 
     * If there has been no previous start of the OOo server, the kill does
     * nothing.
     * 
     * If there has been a previous start, kill destroys the process.
     */
    public void kill() {

        if (oooProcess != null)
        {
            oooProcess.destroy();
            oooProcess = null;
        }
    }

    private static void pipe(final InputStream in, final PrintStream out, final String prefix ) {
       new Thread( "Pipe: " + prefix) {
            @Override
            public void run() {
                BufferedReader r = new BufferedReader(new InputStreamReader(in));
                try {
                    for ( ; ; ) {
                        String s = r.readLine();
                        if (s == null) {
                            break;
                        }
                        out.println(prefix + s);
                    }
                } catch (java.io.IOException e) {
                    e.printStackTrace(System.err);
                }
            }
        }.start();
    }

    /**
     * Returns the list of default options.
     * 
     * @return     The list of default options
     */
    public static List<String> getDefaultOOoOptions() {

        List<String> options = new ArrayList<String>();

        options.add("-nologo");
        options.add("-nodefault");
        options.add("-norestore");
        options.add("-nocrashreport");
        options.add("-nolockcheck");

        return options;
    }
}