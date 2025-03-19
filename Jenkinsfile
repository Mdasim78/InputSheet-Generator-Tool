node {
  stage('SCM') {
    checkout scm
  }
  stage('SonarQube Analysis') {
    def mvn = tool 'Maven';
   withSonarQubeEnv('SonarQube-Server') {
      bat "\"${mvn}\\bin\\mvn\" clean verify sonar:sonar -Dsonar.projectKey=Mdasim78_InputSheet-Generator-Tool_b9a19f5f-78c5-4de8-98c5-078a04cf854a -Dsonar.projectName='InputSheet-Generator-Tool'"
            + "-Dsonar.host.url=http://10.215.123.115:9000/ -Dsonar.login=sqp_3a81bc76cfbe59fc06f8d2dd549510c5e2fac6c9"
   }
  }
}
