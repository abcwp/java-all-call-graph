package test.jacg;

import com.adrninistrator.jacg.other.FindKeywordCallGraph;
import com.adrninistrator.jacg.other.GenSingleCallGraph;

/**
 * @author adrninistrator
 * @date 2021/7/29
 * @description: 生成包含关键字的方法到起始方法之间的调用链，查看方法向下调用链时使用，按照层级增大的方向显示
 */

public class TestFindKeywordCallGraph4er {

    public static void main(String[] args) {
        GenSingleCallGraph.setOrder4er();
        FindKeywordCallGraph findKeywordCallGraph = new FindKeywordCallGraph();
        findKeywordCallGraph.find(args);
    }
}
