package thread.cheaper1.synchronized2;

public class MyThread extends Thread {
    private int i = 5;

    @Override
    public void run() {
        System.out.println(" i=" + (i--) + " threadName=" + Thread.currentThread().getName());
        //注意： 代码 i-- 由 前面 项目 中 单独 一行 运行 改成 在当 前项 目中 在 println() 方法 中 直接进行 打印

        super.run();
    }
}
