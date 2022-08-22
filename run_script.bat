:: Build the Docker image for pears_sites_report.py
docker build -t il_fcs/pears_sites_report:latest .
:: Create and start the Docker container
docker run --name pears_sites_report il_fcs/pears_sites_report:latest
:: Copy /example_outputs from the container to the build context
docker cp pears_sites_report:/pears_sites_report/example_outputs/ ./
:: Remove the container
docker rm pears_sites_report
pause