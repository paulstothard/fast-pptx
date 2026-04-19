FROM node:22-bookworm

# Install system dependencies
RUN apt-get update && apt-get install -y --no-install-recommends \
    graphviz \
    imagemagick \
    pandoc \
    poppler-utils \
    python3 \
    && rm -rf /var/lib/apt/lists/*

# Install global Node.js tools
RUN npm install -g @mermaid-js/mermaid-cli csv2md svgexport

# Copy the full project to a fixed location
COPY . /usr/local/fast-pptx

# Make all scripts executable and readable
RUN chmod -R 755 /usr/local/fast-pptx/scripts && \
    chmod 755 /usr/local/fast-pptx/fast-pptx.sh

# Create python symlink (script expects 'python' not 'python3')
RUN ln -s /usr/bin/python3 /usr/bin/python

# Set Playwright browser cache to a fixed location
ENV PLAYWRIGHT_BROWSERS_PATH=/usr/local/fast-pptx/browsers

# Install Playwright and its Chromium browser (with all Linux deps)
RUN npm install --prefix /usr/local/fast-pptx/scripts \
    && npx --prefix /usr/local/fast-pptx/scripts playwright install chromium --with-deps \
    && chmod -R 755 /usr/local/fast-pptx/browsers

ENTRYPOINT ["/usr/local/fast-pptx/fast-pptx.sh"]
