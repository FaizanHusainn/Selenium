<project xmlns="http://maven.apache.org/POM/4.0.0" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://maven.apache.org/POM/4.0.0 https://maven.apache.org/xsd/maven-4.0.0.xsd">
  <modelVersion>4.0.0</modelVersion>
  <groupId>seleniumwebdriver</groupId>
  <artifactId>seleniumwebdriver</artifactId>
  <version>0.0.1-SNAPSHOT</version>
  
  
  <parent>
    <groupId>org.springframework.boot</groupId>
    <artifactId>spring-boot-starter-parent</artifactId>
    <version>3.1.3</version> <!-- Use the latest stable version here -->
    <relativePath/> <!-- lookup parent from repository -->
</parent>
  
  <dependencies>
<!-- https://mvnrepository.com/artifact/org.seleniumhq.selenium/selenium-java -->
	<dependency>
		 <groupId>org.seleniumhq.selenium</groupId>
		 <artifactId>selenium-java</artifactId>
		 <version>4.23.1</version>
	</dependency>
	
	<dependency>
   <groupId>com.opencsv</groupId>
   <artifactId>opencsv</artifactId>
   <version>5.5.2</version>
  </dependency>
   
   <dependency>
            <groupId>org.springframework.boot</groupId>
            <artifactId>spring-boot-starter-web</artifactId>
        </dependency>
        
        <!-- Thymeleaf Template Engine -->
        <dependency>
            <groupId>org.springframework.boot</groupId>
            <artifactId>spring-boot-starter-thymeleaf</artifactId>
        </dependency>

<dependency>

   <groupId>javax.servlet</groupId>

   <artifactId>javax.servlet-api</artifactId>

   <version>4.0.1</version>

   <scope>provided</scope>

</dependency>
 <!-- Apache POI for working with Word files -->
    <dependency>
        <groupId>org.apache.poi</groupId>
        <artifactId>poi-ooxml</artifactId>
        <version>5.2.3</version>
    </dependency>
    
    

    <!-- Apache POI dependencies for XML schemas and formats -->
    <dependency>
        <groupId>org.apache.poi</groupId>
        <artifactId>poi-ooxml-schemas</artifactId>
        <version>4.1.2</version>
    </dependency>
    
    <!-- XMLBeans (used by POI) -->
    <dependency>
        <groupId>org.apache.xmlbeans</groupId>
        <artifactId>xmlbeans</artifactId>
        <version>5.1.1</version>
    </dependency>

    <!-- Commons Collections (used by POI) -->
    <dependency>
        <groupId>org.apache.commons</groupId>
        <artifactId>commons-collections4</artifactId>
        <version>4.4</version>
    </dependency>
    
      <dependency>
        <groupId>org.springframework.boot</groupId>
        <artifactId>spring-boot-starter-web</artifactId>
    </dependency>

    <!-- Spring Boot DevTools (Optional, for development) -->
    <dependency>
        <groupId>org.springframework.boot</groupId>
        <artifactId>spring-boot-devtools</artifactId>
        <scope>runtime</scope>
        <optional>true</optional>
    </dependency>
    
      <dependency>
        <groupId>org.springframework.boot</groupId>
        <artifactId>spring-boot-starter-thymeleaf</artifactId>
    </dependency>
    
</dependencies>
 <build>
    <plugins>
        <plugin>
            <groupId>org.apache.maven.plugins</groupId>
            <artifactId>maven-compiler-plugin</artifactId>
            <version>3.8.0</version>
            <configuration>
                <source>8</source>
                <target>8</target>
            </configuration>
        </plugin>
        <plugin>
            <groupId>org.apache.maven.plugins</groupId>
            <artifactId>maven-surefire-plugin</artifactId>
            <version>3.0.0-M1</version>
            <configuration>
                <useSystemClassLoader>false</useSystemClassLoader>
                <forkCount>1</forkCount>
                <useFile>false</useFile>
                <skipTests>false</skipTests>
                <testFailureIgnore>true</testFailureIgnore>
                <forkMode>once</forkMode>
                <suiteXmlFiles>
                    <suiteXmlFile>testing.xml</suiteXmlFile>
                </suiteXmlFiles>
            </configuration>
        </plugin>
    </plugins>
</build>
<properties>
    <project.build.sourceEncoding>UTF-8</project.build.sourceEncoding>
</properties>
</project>
 
