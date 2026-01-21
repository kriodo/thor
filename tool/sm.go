package tool

import (
	"bytes"
	"encoding/hex"
	"errors"

	"github.com/tjfoc/gmsm/sm4"
)

const (
	AlgorithmName = "SM4"
	BlockSize     = 16
)

//
// ======================= Padding =======================
//

// PKCS5Padding (SM4 blockSize = 16，等价 PKCS7)
func pkcs5Padding(src []byte) []byte {
	padding := BlockSize - len(src)%BlockSize
	padtext := bytes.Repeat([]byte{byte(padding)}, padding)
	return append(src, padtext...)
}

func pkcs5UnPadding(src []byte) ([]byte, error) {
	length := len(src)
	if length == 0 {
		return nil, errors.New("invalid padding size")
	}
	unpadding := int(src[length-1])
	if unpadding > BlockSize || unpadding == 0 {
		return nil, errors.New("invalid padding")
	}
	return src[:(length - unpadding)], nil
}

//
// ======================= ECB =======================
//

func encryptEcbPadding(key, data []byte) ([]byte, error) {
	block, err := sm4.NewCipher(key)
	if err != nil {
		return nil, err
	}

	data = pkcs5Padding(data)
	dst := make([]byte, len(data))

	for i := 0; i < len(data); i += BlockSize {
		block.Encrypt(dst[i:i+BlockSize], data[i:i+BlockSize])
	}
	return dst, nil
}

func decryptEcbPadding(key, data []byte) ([]byte, error) {
	block, err := sm4.NewCipher(key)
	if err != nil {
		return nil, err
	}

	if len(data)%BlockSize != 0 {
		return nil, errors.New("ciphertext is not a multiple of the block size")
	}

	dst := make([]byte, len(data))
	for i := 0; i < len(data); i += BlockSize {
		block.Decrypt(dst[i:i+BlockSize], data[i:i+BlockSize])
	}
	return pkcs5UnPadding(dst)
}

//
// ======================= Public API =======================
//

// 对应 Java: encryptEcb(String hexKey, String paramStr)
func EncryptEcb(hexKey string, plainText string) (string, error) {
	key, err := hex.DecodeString(hexKey)
	if err != nil {
		return "", err
	}
	if len(key) != 16 {
		return "", errors.New("sm4 key length must be 16 bytes")
	}

	cipherBytes, err := encryptEcbPadding(key, []byte(plainText))
	if err != nil {
		return "", err
	}
	return hex.EncodeToString(cipherBytes), nil
}

// 对应 Java: decryptEcb(String hexKey, String cipherText)
func DecryptEcb(hexKey string, cipherHex string) (string, error) {
	key, err := hex.DecodeString(hexKey)
	if err != nil {
		return "", err
	}

	cipherBytes, err := hex.DecodeString(cipherHex)
	if err != nil {
		return "", err
	}

	plainBytes, err := decryptEcbPadding(key, cipherBytes)
	if err != nil {
		return "", err
	}
	return string(plainBytes), nil
}

// 对应 Java: verifyEcb
func VerifyEcb(hexKey, cipherHex, plainText string) (bool, error) {
	decryptText, err := DecryptEcb(hexKey, cipherHex)
	if err != nil {
		return false, err
	}
	return decryptText == plainText, nil
}
