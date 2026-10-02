# officecorpus adapter: Excelize (github.com/xuri/excelize/v2 v2.11.0)
FROM golang:1.25-trixie@sha256:2c4c60ef415fbfa5e90300722293bef36c5e63fae17570ce18f580af933dbd73 AS build
ENV GOTOOLCHAIN=local CGO_ENABLED=0
WORKDIR /src
COPY go.mod go.sum ./
RUN go mod download && go mod verify
COPY main.go ./
RUN go build -trimpath -o /out/officecorpus-excelize .

FROM debian:trixie-slim@sha256:a99cfc517144bc59b1978475ec53b46ecabec7e43635402ee5b77cc54cd1b20a
LABEL org.officecorpus.adapter="excelize" \
      org.officecorpus.library="github.com/xuri/excelize/v2" \
      org.officecorpus.library-version="2.11.0" \
      org.officecorpus.runtime="Go 1.25 (static, CGO disabled)"
COPY --from=build /out/officecorpus-excelize /usr/local/bin/officecorpus-excelize
USER 65534:65534
ENTRYPOINT ["/usr/local/bin/officecorpus-excelize"]
