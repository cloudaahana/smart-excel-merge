import java.io.BufferedReader;
import java.io.InputStreamReader;

public class PythonHandler {

    public static void executePython() throws Exception {
        String logs;
        System.out.println("\n Calling Python Upload Job | Big Query \n");
        Process pythonProcess = Runtime.getRuntime().exec("python upload_master_csv.py");
        BufferedReader stdInput = new BufferedReader(new InputStreamReader(pythonProcess.getInputStream()));

        BufferedReader stdError = new BufferedReader(new InputStreamReader(pythonProcess.getErrorStream()));

        // read the output from the command
        while ((logs = stdInput.readLine()) != null) {
            System.out.println(logs);
        }
        // read any errors from the attempted command
        while ((logs = stdError.readLine()) != null) {
            System.out.println(logs);
        }
    }
}