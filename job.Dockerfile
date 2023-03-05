FROM ubuntu:latest

RUN apt-get update && \
    apt-get install -y openjdk-17-jdk

WORKDIR /app
COPY ./build/libs/crawling.jar .

ENTRYPOINT ["java", "-jar", "crawling.jar"]
CMD ["1"]