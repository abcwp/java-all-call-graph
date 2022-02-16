package test.jacg;

import com.maker.CallGraphMaker;

public class TestCallGraphMaker {
    public static void main(String[] args) {
        try {
            long startTime = System.currentTimeMillis();
            new CallGraphMaker().run();
            long spendTime = System.currentTimeMillis() - startTime;
            System.out.print("耗时: " + spendTime / 1000.0D);
        } catch (Exception e) {
            System.out.print(e);
        }
    }
}
