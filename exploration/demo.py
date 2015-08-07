def oppNone():
    if numSigs == 1:
      for i in range(0, len(sig1)):
        if sig1[i] == 1:
          array.append(1)
          numTrue += 1
        else:
          array.append(0)
    else:
      print("invalid sigs: none", numSigs)
      return None
def oppAnd():
    if numSigs == 2:
      for i in range(0, len(sig1)):
        if sig1[i] + sig2[i] == 2:
          array.append(1)
          numTrue += 1
        else:
          array.append(0)
    elif numSigs == 3:
      for i in range(0, len(sig1)):
        if sig1[i] + sig2[i] + sig3[i]== 3:
          array.append(1)
          numTrue += 1
        else:
          array.append(0)
    elif numSigs == 4:
      for i in range(0, len(sig1)):
        if sig1[i] + sig2[i] + sig3[i] + sig4[i] == 4:
          array.append(1)
          numTrue += 1
        else:
          array.append(0)
    elif numSigs == 5:
      for i in range(0, len(sig1)):
        if sig1[i] + sig2[i] + sig3[i] + sig4[i] + sig5[i]== 5:
          array.append(1)
          numTrue += 1
        else:
          array.append(0)
    else:
      print("invalid sigs: and", numSigs)
      return None

def oppOr():
    if numSigs == 2:
      for i in range(0, len(sig1)):
        if sig1[i] + sig2[i] >= 1:
          array.append(1)
          numTrue += 1
        else:
          array.append(0)
    elif numSigs == 3:
      for i in range(0, len(sig1)):
        if sig1[i] + sig2[i] + sig3[i] >= 1:
          array.append(1)
          numTrue += 1
        else:
          array.append(0)
    elif numSigs == 4:
      for i in range(0, len(sig1)):
        if sig1[i] + sig2[i] + sig3[i] + sig4[i] >= 1:
          array.append(1)
          numTrue += 1
        else:
          array.append(0)
    elif numSigs == 5:
      for i in range(0, len(sig1)):
        if sig1[i] + sig2[i] + sig3[i] + sig4[i] + sig5[i] >= 1:
          array.append(1)
          numTrue += 1
        else:
          array.append(0)
    else:
      print("invalid sigs: or", numSigs)
      return None
