package thread;

// test
public class Run1 extends Thread {
    public static void main(String[] args) {
        System.out.println(Thread.currentThread().getName());
    }

    public Run1() {
        System.out.println(" 构造 方法 的 打印：" + Thread.currentThread().getName());
    }

    @Override
    public void run() {
        System.out.println(" run 方法 的 打印：" + Thread.currentThread().getName());
    }
}


