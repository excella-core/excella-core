#!/bin/bash

mvn clean jacoco:prepare-agent verify jacoco:report sonar:sonar -Dmaven.javadoc.failOnError=false -s dev/settings-sonarqube.xml $@
