说明：aspose-cells-8.6.2.jar 为excel操作类库，aspose-words-15.8.0-jdk16.jar 为doc文档操作类库
部署：

1，安装jar文件，可以自己上传至maven私服，
      <dependency>
            <groupId>com.aspose</groupId>
            <artifactId>aspose-words</artifactId>
            <version>15.8.0</version>
            <classifier>jdk16</classifier>
        </dependency>


也可以以本地依赖方式部署
   <dependency>
            <groupId>com.aspose</groupId>
            <artifactId>aspose-words</artifactId>
            <version>15.8.0</version>
            <scope>system</scope>
            <systemPath>${project.basedir}/doc/lib/aspose-words-15.8.0-jdk16.jar</systemPath>
            <classifier>jdk16</classifier>
        </dependency>

2，license.xml 放置能读取位置


3，运行代码
