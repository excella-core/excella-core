mvn clean jacoco:prepare-agent verify jacoco:report -Dmaven.javadoc.failOnError=false -s dev\settings-sonarqube.xml %*
