package com.dragon.learnapachepoi.poiword;

/**
 * @author DragonWen
 */
public class ImageEntity {
	
	/**
     * 图片宽度
     */
    private int width = 400;

    /**
     * 图片高度
     */
    private int height = 300;

    /**
     * 图片地址
     */
    private String url;

    /**
     * 图片类型
     * @see ImageUtils.ImageType
     */
    private ImageUtils.ImageType typeId = ImageUtils.ImageType.PNG;
    
    public int getWidth() {
        return width;
    }

    public void setWidth(int width) {
        this.width = width;
    }

    public int getHeight() {
        return height;
    }

    public void setHeight(int height) {
        this.height = height;
    }

    public String getUrl() {
        return url;
    }

    public void setUrl(String url) {
        this.url = url;
    }

    public ImageUtils.ImageType getTypeId() {
        return typeId;
    }

    public void setTypeId(ImageUtils.ImageType typeId) {
        this.typeId = typeId;
    }
    
}
