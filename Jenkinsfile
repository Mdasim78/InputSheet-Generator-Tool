// node {
//   stage('SCM') {
//     checkout scm
//   }
//   stage('SonarQube Analysis') {
//     def mvn = tool 'Maven';
//    withSonarQubeEnv('SonarQube-Server') {
//       bat "\"${mvn}\\bin\\mvn\" clean verify sonar:sonar -Dsonar.projectKey=Mdasim78_InputSheet-Generator-Tool_b9a19f5f-78c5-4de8-98c5-078a04cf854a -Dsonar.projectName='InputSheet-Generator-Tool'"
//             + "-Dsonar.host.url=http://10.215.123.115:9000/ -Dsonar.login=${inputSheetGenerator_sonarqube_token}"
//    }
//   }
// }

node {
  stage('SCM') {
    checkout scm
  }
  stage('SonarQube Analysis') {
    def mvn = tool 'Maven';
    
    withSonarQubeEnv('SonarQube-Server') {
      withCredentials([string(credentialsId: 'inputSheetGenerator_sonarqube_token', variable: 'SONAR_TOKEN')]) {
        bat """
        "${mvn}\\bin\\mvn" clean verify sonar:sonar ^
        -Dsonar.projectKey=Mdasim78_InputSheet-Generator-Tool_b9a19f5f-78c5-4de8-98c5-078a04cf854a ^
        -Dsonar.projectName='InputSheet-Generator-Tool' ^
        -Dsonar.host.url=http://10.215.123.115:9000/ ^
        -Dsonar.token=%SONAR_TOKEN%
        """
      }
    }
  }
}

