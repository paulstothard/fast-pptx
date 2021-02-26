FROM ubuntu:20.04
LABEL maintainer="stothard@ualberta.ca"
LABEL description="Quickly make a PowerPoint presentation from a directory of URLs, images, PDFs, CSV files, and code snippets."

ENV DEBIAN_FRONTEND=noninteractive

RUN apt-get update && apt-get install -y \
  apt-utils \
  build-essential \
  chromium-browser \
  graphviz \
  imagemagick \
  nodejs \
  npm \
  pandoc \
  poppler-utils

RUN npm install -g csv2md

RUN npm install --verbose --global pageres-cli

WORKDIR /usr/bin

COPY scripts/fast-pptx.sh .
