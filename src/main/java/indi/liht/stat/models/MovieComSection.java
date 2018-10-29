package indi.liht.stat.models;

/**
 * Usage:
 * 电影公司+负责的部分
 * @author lihongtao ibraxwell@sina.com
 * on 2018/10/28
 **/
public class MovieComSection {

    /** 电影公司 */
    private String movieCom;

    /** 分工 */
    private String section;

    /**
     * 全参构造函数
     * @param movieCom 电影公司
     * @param section 分工
     */
    public MovieComSection(String movieCom, String section) {
        this.movieCom = movieCom;
        this.section = section;
    }

    /**
     * 获取 电影公司
     * @return 电影公司
     */
    public String getMovieCom() {
        return movieCom;
    }

    /**
     * 设置 电影公司
     * @param movieCom 电影公司
     */
    public void setMovieCom(String movieCom) {
        this.movieCom = movieCom;
    }

    /**
     * 获取 分工
     * @return 分工
     */
    public String getSection() {
        return section;
    }

    /**
     * 设置 分工
     * @param section 分工
     */
    public void setSection(String section) {
        this.section = section;
    }

    @Override
    public String toString() {
        return "MovieComSection{" +
                "movieCom='" + movieCom + '\'' +
                ", section='" + section + '\'' +
                '}';
    }
}
