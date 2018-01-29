"use strict";

const proxyquire = require("proxyquire");

describe("Encryptor", () => {
    let Encryptor, XmlParser, XmlBuilder, encryptor;

    beforeEach(() => {
        Encryptor = proxyquire("../../lib/Encryptor", {
            '@noCallThru': true
        });
        encryptor = new Encryptor();
    });

    // N.B. Don't bother testing most of these. They are extremely complex. We'll get them covered in the E2E tests.

    describe("_buildEncryptionInfo", () => {
        it("should build the XML", () => {
            const input = {
                package: {
                    cipherAlgorithm: 'PKG_CIPHER_ALGORITHM',
                    cipherChaining: 'PKG_CIPHER_CHAINING',
                    saltValue: Buffer.from([1, 2, 3]),
                    hashAlgorithm: 'PKG_HASH_ALGORITHM',
                    hashSize: 12,
                    blockSize: 34,
                    keyBits: 56
                },
                key: {
                    cipherAlgorithm: 'KEY_CIPHER_ALGORITHM',
                    cipherChaining: 'KEY_CIPHER_CHAINING',
                    saltValue: Buffer.from([4, 5, 6]),
                    hashAlgorithm: 'KEY_HASH_ALGORITHM',
                    hashSize: 79,
                    blockSize: 90,
                    spinCount: 21,
                    keyBits: 43,
                    encryptedKeyValue: Buffer.from([6, 5, 4]),
                    encryptedVerifierHashInput: Buffer.from([3, 2, 1]),
                    encryptedVerifierHashValue: Buffer.from([0, 1, 2])
                },
                dataIntegrity: {
                    encryptedHmacKey: Buffer.from([7, 8, 9]),
                    encryptedHmacValue: Buffer.from([9, 8, 7])
                }
            };

            const output = encryptor._buildEncryptionInfo(input);
            expect(output.slice(0, 8)).toEqualUInt8Array(Buffer.from([0x04, 0x00, 0x04, 0x00, 0x40, 0x00, 0x00, 0x00]));
            expect(output.slice(8).toString()).toBe(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password" xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate"><keyData saltSize="3" blockSize="34" keyBits="56" hashSize="12" cipherAlgorithm="PKG_CIPHER_ALGORITHM" cipherChaining="PKG_CIPHER_CHAINING" hashAlgorithm="PKG_HASH_ALGORITHM" saltValue="AQID"/><dataIntegrity encryptedHmacKey="BwgJ" encryptedHmacValue="CQgH"/><keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password"><p:encryptedKey spinCount="21" saltSize="3" blockSize="90" keyBits="43" hashSize="79" cipherAlgorithm="KEY_CIPHER_ALGORITHM" cipherChaining="KEY_CIPHER_CHAINING" hashAlgorithm="KEY_HASH_ALGORITHM" saltValue="BAUG" encryptedVerifierHashInput="AwIB" encryptedVerifierHashValue="AAEC" encryptedKeyValue="BgUE"/></keyEncryptor></keyEncryptors></encryption>`);
        });
    });

    describe("_parseEncryptionInfoAsync", () => {
        itAsync("should parse the encryption info", () => {
            const xml = `<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password" xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate"><keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="UWYgcVRmEQ/aHrvzqA7xnQ=="/><dataIntegrity encryptedHmacKey="9On0eyZfGVgmwHm4Fi1tV2640oW2wKPEJrU4UY/FUuS2uYh2sh5GRn2mvZ9ifaCOI0P8kdtVcaqDkvxOWODVrw==" encryptedHmacValue="NHg7giBi9SGaKJV3dq4seA+dFaaTYJNkuDBWI0ct92hhJ8mqvzwfiUAyo5a/f+fUmP7QdtH4LIADvgGKXiJLEw=="/><keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password"><p:encryptedKey spinCount="100000" saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="hcJBsEzDpOwkH2qzS9eo3Q==" encryptedVerifierHashInput="T0cM1hi2GTWxzwEa0zZ4vg==" encryptedVerifierHashValue="Pz9v8OrlVkcbrcdfDrxjzD92phbLdUGgifErOSx84RD3E9/c52bVYea9gK+luia2DR727ecXkAjqJT6KGpaOMw==" encryptedKeyValue="hQZ/4Gzp34ILXQ0zc/pRe3JjZVoAjAl2cl1hA56ww9E="/></keyEncryptor></keyEncryptors></encryption>`;
            const input = Buffer.concat([Buffer.alloc(8), Buffer.from(xml)]);
            return encryptor._parseEncryptionInfoAsync(Buffer.from(input))
                .then(output => {
                    expect(output).toEqualJson({
                        package: {
                            cipherAlgorithm: 'AES',
                            cipherChaining: 'ChainingModeCBC',
                            saltValue: Buffer.from([0x51, 0x66, 0x20, 0x71, 0x54, 0x66, 0x11, 0x0f, 0xda, 0x1e, 0xbb, 0xf3, 0xa8, 0x0e, 0xf1, 0x9d]),
                            hashAlgorithm: 'SHA512',
                            blockSize: 16
                        },
                        key: {
                            encryptedKeyValue: Buffer.from([0x85, 0x06, 0x7f, 0xe0, 0x6c, 0xe9, 0xdf, 0x82, 0x0b, 0x5d, 0x0d, 0x33, 0x73, 0xfa, 0x51, 0x7b, 0x72, 0x63, 0x65, 0x5a, 0x00, 0x8c, 0x09, 0x76, 0x72, 0x5d, 0x61, 0x03, 0x9e, 0xb0, 0xc3, 0xd1]),
                            cipherAlgorithm: 'AES',
                            cipherChaining: 'ChainingModeCBC',
                            saltValue: Buffer.from([0x85, 0xc2, 0x41, 0xb0, 0x4c, 0xc3, 0xa4, 0xec, 0x24, 0x1f, 0x6a, 0xb3, 0x4b, 0xd7, 0xa8, 0xdd]),
                            hashAlgorithm: 'SHA512',
                            spinCount: 100000,
                            keyBits: 256
                        }
                    });
                });
        });
    });

    describe("_hash", () => {
        it("should calculate the hash", () => {
            const output = encryptor._hash("SHA256", Buffer.from([0x66, 0x6f, 0x6f]), Buffer.from([0x62, 0x61, 0x72]))
            expect(output.toString("base64")).toBe("w6uP8Tcg6K2QR905Rms8iXTlksL6OD1KOWBxTK7wxPI=");
        });
    });

    describe("_hmac", () => {
        it("should calculate the hash", () => {
            const output = encryptor._hash("SHA256", Buffer.from([1, 2, 3, 4, 5]), Buffer.from([0x66, 0x6f, 0x6f]), Buffer.from([0x62, 0x61, 0x72]))
            expect(output.toString("base64")).toBe("H2SZlE1Npwp6Uw7Ee061LK6veJVKmj6SXO7ldVgTpa8=");
        });
    });

    describe("_crypt", () => {
        const key = Buffer.from([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16]);
        const iv = Buffer.from([16, 15, 14, 13, 12, 11, 10, 9, 8, 7, 6, 5, 4, 3, 2, 1]);
        const plainText = Buffer.from("This is top secret! This is top secret! This is top secret! This is top secret! "); // Must be multiple of block length, which this happens to be"
        const cipherText = Buffer.from("KEPYtNg0JfAy6H20RggidB7s6IiA16lJYrL/MfushFSGCwu6g1nWE+kbOe2/LnaNXPRtnXmFhhC/ITLEPFuWkpN8nT6FjDht2NCzmnzp85E=", "base64");

        it("should encrypt the data", () => {
            const output = encryptor._crypt( true, "AES", "ChainingModeCBC", key, iv, plainText);
            expect(output).toEqualUInt8Array(cipherText)
        });

        it("should decrypt the data", () => {
            const output = encryptor._crypt(false, "AES", "ChainingModeCBC", key, iv, cipherText);
            expect(output).toEqualUInt8Array(plainText);
        });
    });

    describe("_createUInt32LEBuffer", () => {
        it("should create a 4 byte buffer by default", () => {
            const output = encryptor._createUInt32LEBuffer(1234);
            expect(output).toEqualUInt8Array(Buffer.from([210, 4, 0, 0]));
        });

        it("should create a buffer of given length", () => {
            const output = encryptor._createUInt32LEBuffer(4321, 7);
            expect(output).toEqualUInt8Array(Buffer.from([225, 16, 0, 0, 0, 0, 0]));
        });
    });
});
