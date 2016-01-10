package org.bbreak.excella.core;

import java.io.File;

/**
 * テスト用のユーティリティクラス
 * 
 * @since 1.10
 */
public class CoreTestUtil {

    /**
     * テスト用出力ディレクトリを作成する
     * 
     * @return
     */
    public static String getTestOutputDir() {

        String tempDir = System.getProperty( "user.dir") + "/work/test/";
        File file = new File( tempDir);
        if ( !file.exists()) {
            file.mkdirs();
        }

        return tempDir;
    }
}
