mvn clean package -T 1C -Dmaven.skip.test=true
docker stop compare-cells-java
docker rm compare-cells-java
docker rmi compare-cells-java:v1
docker build -t compare-cells-java:v1 .
docker run -d --name compare-cells-java -p 2023:2023 compare-cells-java:v1
