import fs from 'fs';
import {publicKeyPath,privateKeyPath} from '../../config'
import jwt from "jsonwebtoken"

class hldAuthen{

    static async getToken(ctx) {
        let param = ctx.request.body;
        param.expiresAt = 30;
        let publicKey = fs.readFileSync(publicKeyPath);
        let token = jwt.sign(param, publicKey, { algorithm: "RS256", expiresIn: '30m' });
        let data = { result: { status: 1, token: token } };
        ctx.body = data;
        return;
    }

    static async verifyToken(ctx) {
        const param = ctx.request.body;
        let token = param.token;
        let privateKey = fs.readFileSync(privateKeyPath);
        let jwtData = jwt.verify(token, privateKey);
        ctx.body = jwtData;
        return;
    }
}

module.exports = hldAuthen;