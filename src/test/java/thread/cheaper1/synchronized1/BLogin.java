package thread.cheaper1.synchronized1;

public class BLogin extends Thread {
    @Override
    public void run() {
        LoginServlet.doPost("b","bb");
        super.run();
    }
}
