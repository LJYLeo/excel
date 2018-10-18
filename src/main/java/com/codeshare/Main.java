package com.codeshare;

import com.codeshare.config.Config;
import com.codeshare.utils.ExcelUtils;

public class Main {

    public static void main(String[] args) {

        Config.loadVillageConfig();

        ExcelUtils.process();

    }

}
