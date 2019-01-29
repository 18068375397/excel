package thread.cheaper1.synchronized1;

public class ALogin extends Thread {
    @Override
    public void run() {
        LoginServlet.doPost("a","aa");
        super.run();
    }
}
