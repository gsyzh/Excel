package cn.excel.demo.bean;

import lombok.Data;

import java.math.BigDecimal;

/**
 * @author gsyzh
 */
@Data
public class Electric {
    private String type;
    private BigDecimal volume;
    private BigDecimal actualVolume;
    private BigDecimal price;
    private BigDecimal cost;
    private BigDecimal surplusVolume;
    private int index;

    public Electric(String type, BigDecimal volume, BigDecimal actualVolume, BigDecimal price, BigDecimal cost, BigDecimal surplusVolume,int index) {
        this.type = type;
        this.volume = volume;
        this.actualVolume = actualVolume;
        this.price = price;
        this.cost = cost;
        this.surplusVolume = surplusVolume;
        this.index = index;
    }
}
