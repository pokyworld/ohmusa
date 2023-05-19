var c_Null = " can not be null!"
var c_Delete = "Do you really want to delete this record?"
var c_InvalidEmail = "Invalid email, please try again!"
function isBlank(p_Str) {
  var v_Len; if (p_Str == null)
    v_Len = 0; else
    v_Len = p_Str.length; if (!v_Len)
    return true; for (i = 0; i < v_Len; i++) {
    if (p_Str.charAt(i) != " ")
      return false;
  }
  return true;
}
function isEmail(p_Str) {
  var v_Len; if (p_Str == null)
    v_Len = 0; else
    v_Len = p_Str.length; if (!v_Len)
    return false; if (p_Str.indexOf(' ') >= 0)
    return false; if (p_Str.indexOf('@') == -1)
    return false; if (p_Str.indexOf('.') == -1)
    return false; return true;
}
function isPassValid(passwordStr) {
  if (passwordStr == "") { alert("Please enter a valid password...!"); return false; }
  if (passwordStr.length < 6) { alert("Password should not be less than 6 characters...!"); return false; }
  if (passwordStr.length > 12) { alert("Password should not be greater than 12 characters...!"); return false; }
  return true;
}
function isPassValid2() {
  if (1 == 1) { alert("passwordStr2"); return false; }
  return true;
}
function isDayOMOY(vDay, vMonth, vYear) {
  if (parseInt(vYear) < 1900)
    return false; if ((vDay < 1) || (vMonth < 1) || (vMonth > 12) || (vYear < 0))
    return false; if ((vMonth == 1) || (vMonth == 3) || (vMonth == 5) || (vMonth == 7) || (vMonth == 8) || (vMonth == 10) || (vMonth == 12))
    if (vDay > 31)
      return false; if ((vMonth == 4) || (vMonth == 6) || (vMonth == 9) || (vMonth == 11))
    if (vDay > 30)
      return false; if (vMonth == 2)
    if (vYear % 4 == 0)
      return (vDay <= 29); else
      return (vDay <= 28); return true;
}
function isDate(str, mode) {
  var vlen, i, vType, aDate; vType = "/"; aDate = str.split(vType); var alen; alen = aDate.length; if (alen != 3)
    return false; if ((aDate[0].length != 1) && (aDate[0].length != 2))
    return false; if ((aDate[1].length != 1) && (aDate[1].length != 2))
    return false; if ((aDate[2].length != 2) && (aDate[2].length != 4))
    return false; for (i = 0; i < alen; i++) {
      if (isNaN(aDate[i]))
        return false;
    }
  if (mode == 1)
    return (isDayOMOY(aDate[0], aDate[1], aDate[2])); else
    return (isDayOMOY(aDate[1], aDate[0], aDate[2])); return true;
}
function checkyear(vld) {
  var str = vld
  var str1 = "/"
  var s = str.lastIndexOf(str1)
  aa = str.substring(s + 1, str.length)
  if (aa < 1900) { return false; } else
    return true;
}
function delOnClick() {
  var choose; choose = confirm("Do you really want to delete this record?"); if (!choose) { return false; }
  return true;
}
function isInt(num) {
  var tempNum = new String(num); tempNum = allTrim(tempNum); var numLen = tempNum.length; if (numLen == 0) { return true; }
  else {
    var i = 0; for (i = 0; i < numLen; i++) {
      if ((tempNum.substr(i, 1) < "0") || (tempNum.substr(i, 1) > "9")) { return false; }
    }
    return true;
  }
}
function isLength(str, l) {
  if (str.length > l) { return false; }
  return true;
}
function lTrim(str) {
  var strg, vlen; strg = str; vlen = strg.length; while ((vlen > 0) && (strg.charAt(0) == " ")) { strg = strg.substr(1); vlen = strg.length; }
  return (strg);
}
function rTrim(str) {
  var strg, vlen; strg = str; vlen = strg.length; while ((vlen > 0) && (strg.charAt(vlen - 1) == " ")) { strg = strg.substr(0, vlen - 1); vlen = strg.length; }
  return (strg);
}
function allTrim(str) { var strg; strg = lTrim(str); strg = rTrim(strg); return (strg); }
function confirmDelete() {
  var ok = window.confirm(c_Delete); if (ok) { return true; }
  else { return false; }
}
function scrollit_r2l(seed) {
  var msg = "American Technologis-inc very happy service customer!"; var out = " "; if (seed <= 100 && seed > 0) { for (c = 0; c < seed; c++)out += " "; out += msg; seed--; var cmd = "scrollit_r2l(" + seed + ")"; window.status = out; timerTwo = window.setTimeout(cmd, 100); } else
    if (seed <= 0) {
      if (-seed < msg.length) { out += msg.substring(-seed, msg.length); seed--; var cmd = "scrollit_r2l(" + seed + ")"; window.status = out; timerTwo = window.setTimeout(cmd, 100); }
      else { window.status = " "; timerTwo = window.setTimeout("scrollit_r2l(100)", 75) }
    }
}