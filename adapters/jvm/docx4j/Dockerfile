# officecorpus adapter: docx4j (org.docx4j:docx4j-JAXB-ReferenceImpl 17.2.1)
FROM maven:3.9.16-eclipse-temurin-21 AS build
WORKDIR /build
COPY pom.xml .
RUN mvn -B -q dependency:resolve
COPY src ./src
RUN mvn -B -q package

FROM eclipse-temurin:21.0.12_8-jre
LABEL org.officecorpus.adapter="docx4j" \
      org.officecorpus.library="org.docx4j:docx4j-JAXB-ReferenceImpl" \
      org.officecorpus.library-version="17.2.1" \
      org.officecorpus.runtime="OpenJDK 21.0.12+8"
WORKDIR /app
COPY --from=build /build/target/lib/ /app/lib/
COPY --from=build /build/target/officecorpus-docx4j.jar /app/officecorpus-docx4j.jar
USER 65534:65534
# slf4j-simple writes to stderr; stdout carries only the protocol.
ENTRYPOINT ["java", "-Djava.awt.headless=true", "-Dorg.slf4j.simpleLogger.defaultLogLevel=warn", "-cp", "/app/officecorpus-docx4j.jar:/app/lib/*", "org.officecorpus.docx4j.OfficeCorpusDocx4j"]
