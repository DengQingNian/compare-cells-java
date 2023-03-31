FROM openjdk:17-jdk-alpine
WORKDIR /
COPY target/compare-cells-0.0.1-SNAPSHOT.jar /

EXPOSE 2023

CMD ["java", "-jar", "compare-cells-0.0.1-SNAPSHOT.jar"]