# officecorpus adapter: Apache POI usermodel (org.apache.poi:poi-ooxml 5.5.1)
FROM maven:3.9.16-eclipse-temurin-21 AS build
WORKDIR /build
COPY pom.xml .
RUN mvn -B -q dependency:resolve
COPY src ./src
RUN mvn -B -q package

FROM eclipse-temurin:21.0.12_8-jre
LABEL org.officecorpus.adapter="apache-poi" \
      org.officecorpus.library="org.apache.poi:poi-ooxml" \
      org.officecorpus.library-version="5.5.1" \
      org.officecorpus.runtime="OpenJDK 21.0.12+8"
WORKDIR /app
COPY --from=build /build/target/lib/ /app/lib/
COPY --from=build /build/target/officecorpus-apache-poi.jar /app/officecorpus-apache-poi.jar
USER 65534:65534
# log4j-api without an implementation: use the SimpleLogger, which writes to stderr.
ENTRYPOINT ["java", "-Djava.awt.headless=true", "-Dlog4j2.loggerContextFactory=org.apache.logging.log4j.simple.SimpleLoggerContextFactory", "-Dlog4j2.simplelogLevel=WARN", "-cp", "/app/officecorpus-apache-poi.jar:/app/lib/*", "org.officecorpus.poi.OfficeCorpusPoi"]
