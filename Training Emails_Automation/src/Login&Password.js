function loginAndPassword(candidateName) {
  let loginId, password;
  let firstName, lastName;
  if (candidateName != '') {
    // console.log(candidateName);
    const name = candidateName.split(" ");
    if (name.length === 1) {
      firstName = name[0];
      loginId = firstName.substring(0, 8).toLowerCase();
      password = loginId.charAt(0).toUpperCase() + loginId.slice(1) + '@UpThink1';
    } else {
      firstName = name[0];
      lastName = name[name.length - 1];
      let loginIdPart1, loginIdPart2;
      if (lastName.length >= 4) {
        loginIdPart1 = firstName.substring(0, 4).toLowerCase();
        loginIdPart2 = lastName.substring(0, 8 - loginIdPart1.length).toLowerCase();
      } else {
        loginIdPart2 = lastName.substring(0, 4).toLowerCase();
        loginIdPart1 = firstName.substring(0, 8 - loginIdPart2.length).toLowerCase();
      }
      loginId = loginIdPart1 + loginIdPart2;
      password = firstName.charAt(0).toUpperCase() + firstName.slice(1) + lastName.charAt(0).toUpperCase() + '@UpThink1';
        }
  } else {
      loginId = '';
      password = '';
  }

  return {
    "login": loginId,
    "password": password
  }

}
