node {
  stage('SCM') {
    checkout scm
  }
  stage('SonarQube Analysis') {
    def mvn = tool 'Maven';
   withSonarQubeEnv('SonarQube-Server') {
      sh "${mvn}/bin/mvn clean verify sonar:sonar -Dsonar.projectKey=Mdasim78_InputSheet-Generator-Tool_b9a19f5f-78c5-4de8-98c5-078a04cf854a -Dsonar.projectName='InputSheet-Generator-Tool'"
    }
  }
}
