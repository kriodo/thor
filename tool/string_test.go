package tool

import (
	"bytes"
	"encoding/hex"
	"testing"
)

// PKCS7 补位
func pkcs7Padding(data []byte, blockSize int) []byte {
	padding := blockSize - len(data)%blockSize
	padtext := bytes.Repeat([]byte{byte(padding)}, padding)
	return append(data, padtext...)
}

// 去补位
func pkcs7UnPadding(data []byte) []byte {
	length := len(data)
	unpadding := int(data[length-1])
	return data[:(length - unpadding)]
}

func TestUniqueString(t *testing.T) {

	//key := []byte("1234567890abcdef") // 16 字节
	//key, err := getKey() // 16 字节
	//if err != nil {
	//	t.Log(err)
	//	return
	//}
	secret := "6b42698fc0def7b3f664c7e55d3d80556c483b93336ada0b190737240663b741"
	k := secret[0:32]
	ecb, err := EncryptEcb(k, "{\"businessType\":0}")
	if err != nil {
		return
	}
	t.Log(ecb)
	//plaintext := []byte("{\"SYCOMRETZ\":[{\"ERRMSG\":\"请求解密失败，密文格式不对\",\"ERRCOD\":\"OGAPP44\",\"ERRDTL\":\"\",\"ERRPAM\":\"\"}]}")

	//// ---------- 加密 ----------
	//inData := pkcs7Padding(plaintext, 16)
	//
	//ciphertext, err := sm4.Sm4Ecb(key, inData, true) // true = 加密
	//if err != nil {
	//	t.Log(err)
	//	return
	//}
	//b64 := base64.StdEncoding.EncodeToString(ciphertext)
	//fmt.Println("ECB 加密 Base64:", b64)

	// ---------- 解密 ----------
	//decoded, _ := base64.StdEncoding.DecodeString(b64)
	//decoded = []byte("a253e26a86ea25142922d1783294fb3acdbf94f0aeb0f20ae68936b65b396131")
	//decrypted, err := sm4.Sm4Ecb(key, decoded, false) // false = 解密
	//if err != nil {
	//	t.Log(err)
	//	return
	//}
	//result := pkcs7UnPadding(decrypted)

	//fmt.Println("解密明文:", string(result))
}

func getKey() ([]byte, error) {
	secret := "6b42698fc0def7b3f664c7e55d3d80556c483b93336ada0b190737240663b741"
	k := secret[0:32]
	key, err := hex.DecodeString(k)
	if err != nil {
		return nil, err
	}
	return key, nil
}

//7afbe4810244410d6a5903de9260d864e1c5e4a70df4d2d219c00aebb641067a32519189c9b62d3cd6161edb23eaf3d2bc70eb1ea01e9a039a519f6544c396b89076f8c552a9262ca53df7f82c2d308d8f4dabad1a04f75ce54fbd3a9b53eca0f045b39aef962d376618f72530a86eb7
//evvkgQJEQQ1qWQPekmDYZOHF5KcN9NLSGcAK67ZBBnoyUZGJybYtPNYWHtsj6vPSvHDrHqAemgOaUZ9lRMOWuJB2+MVSqSYspT33+CwtMI2PTautGgT3XOVPvTqbU+yg8EWzmu+WLTdmGPclMKhut04BTMVynR4Q+7TDT+iyu2o=
