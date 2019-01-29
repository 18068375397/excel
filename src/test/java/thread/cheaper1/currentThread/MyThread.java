package thread.cheaper1.currentThread;

public class MyThread extends Thread{
    public MyThread() {
        super();
        System.out.println("构造方法打印："+ Thread.currentThread().getName());
    }

    @Override
    public void run() {
        super.run();
        System.out.println("run方法打印："+ Thread.currentThread().getName());
    }
}
