package thread.cheaper1.currentThread;

public class TestThread {

    public static void main(String[] args) {
        new Thread(new Runnable() {
            @Override
            public void run() {
                while (true) {
                    try {
                        new Thread().sleep(500);
                        System.out.println("--runnable->" + Thread.currentThread().getName());
                    } catch (InterruptedException e) {
                        e.printStackTrace();
                    }
                }
            }})
        {
            @Override
            public void run() {
				super.run();
                while (true) {
                    try {
                        new Thread().sleep(500);
                        System.out.println("--thread.run->" + Thread.currentThread().getName());
                    } catch (InterruptedException e) {
                        e.printStackTrace();
                    }
                }
            }
        }.start();
    }
}
