const MainModel = require('../models/m_main');
const { response } = require('../utils/response.utils');
const readXlsxFile = require('read-excel-file/node');
const bcrypt = require('bcrypt');
const jwt = require('jsonwebtoken');
const nodemailer = require('nodemailer');
const _ = require('lodash');
const {logger} = require('../config/winston');
const excel = require("exceljs");
const ejs = require("ejs");
const pdf = require("html-pdf");
const path = require("path");
const dotenv = require('dotenv');
const { concat } = require('lodash');
dotenv.config();

function makeRandom(n) {
	let result = '';
    const characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    const charactersLength = characters.length;
    for ( let i = 0; i < n; i++ ) {
      result += characters.charAt(Math.floor(Math.random() * charactersLength));
   	}
   	return result;
}

function UpperFirstLetter(str) {
	return str.split(' ').map(i => i[0].toUpperCase() + i.substring(1).toLowerCase()).join(' ')
}

function dateconvert(str) {
	const bulan = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Agustus', 'September', 'Oktober', 'November', 'Desember']
	const date = new Date(str);
    const mnth = bulan[date.getMonth()];
    const day = ("0" + date.getDate()).slice(-2);
  	const valueConvert = [day, mnth, date.getFullYear()].join(" ")
	return valueConvert
}

function convertDate(str) {
	let date = new Date(str),
    mnth = ("0" + (date.getMonth() + 1)).slice(-2),
    day = ("0" + date.getDate()).slice(-2);
  	const valueConvert = [date.getFullYear(), mnth, day].join("-");
	return valueConvert
}

class MainController {
  readDataBy = async (attributes = [], params = {}, include = {}, on = null, table) => {
    let defaultWhere = { or: {}, and: {} }
    params = (Object.entries(params).length > 0) ? params : defaultWhere
    let defaultInclude = {}
    include = (Object.entries(include).length > 0) ? include : defaultInclude
    let prototype = {
      attributes: attributes, //kalau null all item selected
      table: table,
      on: on,
      include: include,
      // {
      //   attributes: ['createdAt AS cAtt', 'updatedAt AS uAtt'],
      //   joinTable: 'users_details',
      //   joinOn: 'users_details.id_profile',
      // },
      where: params,
      orderBy: 'ASC',
      orderByValue: '' //default '' => item_no
    }
    // console.log(prototype)
    // process.exit()
    let data = await MainModel.findWhere(prototype);
    return data
  }

  dataDashboard = async (req, res, next) => {
    try{
      // let query = { ...req.query }
      let { role, id_profile } = req.query;
      let attributes = []
      let include = {}
      let search = {}
      let data = await this.readDataBy(attributes, search, include, 'viewusers.id', 'viewusers')
      let dataDashboard
      if(role == 'admin') {
        //Untuk Administrator
        let dataSiswaPria = data.filter((el) => el.roleID === 3 && el.activeAkun === 1 && el.mutationAkun === 0 && el.jeniskelamin === "Laki - Laki").length
        let dataSiswaWanita = data.filter((el) => el.roleID === 3 && el.activeAkun === 1 && el.mutationAkun === 0 && el.jeniskelamin === "Perempuan").length
        let dataSiswaMutasi = data.filter((el) => el.roleID === 3 && (el.activeAkun === 1 || el.activeAkun === 0) && el.mutationAkun === 1).length
        let dataGuru = data.filter((el) => el.roleID === 2 && el.activeAkun === 1).length
        dataDashboard = {
          dataSiswaPria: dataSiswaPria,
          dataSiswaWanita: dataSiswaWanita,
          dataSiswaMutasi: dataSiswaMutasi,
          dataGuru: dataGuru
        }
      } else if(role == 'guru') {
        //Untuk Guru Perseorangan
        let mencariGuru = data.filter((el) => el.id_profile === parseInt(id_profile))
        // console.log(mencariGuru)
        const mengajarKelas = String(mencariGuru[0].mengajar_kelas)
        let MengajarKelas = mengajarKelas.split(', ').sort()
        let jumlahSiswa
        let hasilPush = new Array()
        MengajarKelas.map((kelas) => {
          jumlahSiswa = data.filter((el) => el.roleID === 3 && el.activeAkun === 1 && el.mutationAkun === 0 && el.kelas === kelas).length
          hasilPush.push({kelas, jumlahSiswa})
        })
        dataDashboard = {
          dataGuru: mencariGuru[0],
          total: hasilPush
        }
      } else if(role == 'siswa') {
        dataDashboard = null
      }
      let kode = 200;
      let message = 'Berhasil';
      response(res, { kode, message, result: dataDashboard }, 200);
    }catch(err){
      console.log(err)
      logger.error(`[ERROR ACCESS] --- Gagal memproses data !`);
      return response(res, { kode: 404, message: 'Gagal memproses data !' }, 404);
    }
  }

  login = async (req, res, next) => {
    try{
      let { username, password, jenis, gambar } = req.body
      let attributes = []
      let include = {}
      let search = (jenis == 'nonGmail') ? 
        {
          or: {
            email: username, 
            nomor_induk: username,
          },
          and: {
            activeAkun: 1
          }
        } :
        {
          or: {},
          and: {
            email: username, 
            activeAkun: 1
          }
        }

      let data = await this.readDataBy(attributes, search, include, 'viewusers.id', 'viewusers');
      if(!data.length){ return response(res, { kode: 404, message: 'Data tidak di temukan !' }, 404); }

      if(jenis == 'nonGmail'){
        const match = await bcrypt.compare(password, data[0].password);
        if(!match) return response(res, { kode: 404, message: 'Kata Sandi tidak sesuai !' }, 404);
      }
      const userID = data[0].id;
      const name = data[0].name;
      const email = data[0].email;
      const accessToken = jwt.sign({userID, name, email}, process.env.ACCESS_TOKEN_SECRET, {
          expiresIn: '12h'
      });
      const refreshToken = jwt.sign({userID, name, email}, process.env.REFRESH_TOKEN_SECRET, {
          expiresIn: '1d'
      });

      const kirimData ={
        refresh_token: refreshToken,
        codeLog: (jenis == 'nonGmail') ? '1' : '2',
        gambarGmail: (jenis == 'nonGmail') ? null : gambar
      }

      await MainModel.update(kirimData, 'users', { id: data[0].id });

      res.cookie('refreshToken', refreshToken, {
        httpOnly: true,
        maxAge: 24 * 60 * 60 * 1000,
        // secure: true
      });

      let kode = 200;
      let message = `Anda berhasil masuk panel Dashboard melalui ${(jenis == 'nonGmail')?'Login Panel':'Akun Gmail'} !`;
      logger.info(`[USER ${(data[0].roleID == 1)?'ADMIN':(data[0].roleID == 2)?'GURU':'SISWA/I'} - ${data[0].name}] --- berhasil masuk panel Dashboard melalui ${(jenis == 'nonGmail')?'Login Panel':'Akun Gmail'} !`);
      response(res, { kode, message, result: {...data[0], accessToken} }, 200);
    }catch(err){
      console.log(err)
      logger.error(`[ERROR ACCESS] --- Gagal memproses data !`);
      return response(res, { kode: 404, message: 'Gagal memproses data !' }, 404);
    }
  }

  logout = async (req, res, next) => {
    try{
      let id = req.params.id;
      let attributes = []
      let include = {}
      let search = { or: {}, and: { id: id } }

      let data = await this.readDataBy(attributes, search, include, 'viewusers.id', 'viewusers');
      if(!data.length) return response(res, { kode: 404, message: 'Data tidak di temukan !' }, 404);

      const kirimData ={
        refresh_token: null,
        codeLog: '0',
        gambarGmail: null
      }
      await MainModel.update(kirimData, 'users', { email: data[0].email });
      res.clearCookie('refreshToken');

      let kode = 200;
      let message = 'Anda berhasil keluar panel Dashboard !';
      logger.info(`[USER ${(data[0].roleID == 1)?'ADMIN':(data[0].roleID == 2)?'GURU':'SISWA/I'} - ${data[0].name}] --- berhasil keluar panel Dashboard !`);
      response(res, { kode, message }, 200);
    }catch(err){
      logger.error(`[ERROR ACCESS] --- Gagal memproses data !`);
      return response(res, { kode: 404, message: 'Gagal memproses data !' }, 404);
    }
  }

  readAllData = async (req, res, next) => {
    try{
      let roleID = req.params.idRole;
      let attributes = []
      let include = {}
      let search = { or: {}, and: { roleID: roleID } }

      let data = await this.readDataBy(attributes, search, include, 'viewusers.id', 'viewusers');
      let kode = 200;
      let message = 'Berhasil';
      logger.info(`[FETCH DATA] --- Load data users by role ID ${(roleID == 1)?'ADMIN':(roleID == 2)?'GURU':'SISWA/I'}`);
      response(res, { kode, message, result: data }, 200);
    }catch(err){
      logger.error(`[ERROR ACCESS] --- Gagal memproses data !`);
      return response(res, { kode: 404, message: 'Gagal memproses data !' }, 404);
    }
  }

  readDataByID = async (req, res, next) => {
    try{
      let id = req.params.id;
      let attributes = []
      let include = {}
      let search = { or: {}, and: { id: id } }

      let data = await this.readDataBy(attributes, search, include, 'viewusers.id', 'viewusers');
      let kode = 200;
      let message = 'Berhasil';
      logger.info(`[FETCH DATA] --- Load data users by ID ${id} - ${data.name}`);
      response(res, { kode, message, result: (typeof data[0] == 'undefined') ? {} : data[0] }, 200);
    }catch(err){
      logger.error(`[ERROR ACCESS] --- Gagal memproses data !`);
      return response(res, { kode: 404, message: 'Gagal memproses data !' }, 404);
    }
  }  

  readDataByIDLookscreen = async (req, res, next) => {
    try{
      let id = req.params.id;
      let attributes = []
      let include = {}
      let search = { or: {}, and: { id: id } }

      let data = await this.readDataBy(attributes, search, include, 'viewusers.id', 'viewusers');
      let kode = 200;
      let message = 'Berhasil';
      logger.info(`[FETCH DATA] --- Load data users by ID ${id} - ${data.name}`);
      response(res, { kode, message, result: (typeof data[0] == 'undefined') ? {} : data[0] }, 200);
    }catch(err){
      logger.error(`[ERROR ACCESS] --- Gagal memproses data !`);
      return response(res, { kode: 404, message: 'Gagal memproses data !' }, 404);
    }
  }  

  createupdateData = async (req, res, next) => {
    try{
      let body = { ...req.body };      
      let attributes = []
      let include = {}
      let search = {
        or: {
          email: body.email,
          nomor_induk: body.nomor_induk,
        },
        and: {
          activeAkun: 1
        },
      }

      let readData = await this.readDataBy(attributes, search, include, 'viewusers.id', 'viewusers');
      let salt, hashPassword, kirimdata1, kirimdata2, transporter, mailOptions, html;
      switch (body.jenis) {
        case ('ADD'):
          if(readData.length){ return response(res, { kode: 404, message: 'Email atau Nomor Induk sudah ada !' }, 404); }
          salt = await bcrypt.genSalt();
          hashPassword = await bcrypt.hash(body.password, salt);
          kirimdata1 = {
            roleID: body.roleID,
            name: body.name,
            email: body.email,
            password: hashPassword,
            kodeOTP: body.password,
          }
          let insertUsers = await MainModel.insert(kirimdata1, 'users');
          if(!insertUsers) { await MainModel.delete({ email: body.email }, 'users'); }
          let search2 = {
            or: {},
            and: {
              email: body.email
            },
          }
          let readData2 = await this.readDataBy(search2, 'users')

          kirimdata2 = {
            id_profile: readData2[0].id,
            nomor_induk: body.roleID == '1' ? '' : body.nomor_induk,
            nik_siswa: body.roleID == '3' ? body.nik_siswa : null,
            tempat: body.tempat,
            tgl_lahir: body.tgl_lahir,
            jeniskelamin: body.jeniskelamin,
            agama: body.agama,
            telp: body.telp,
            alamat: body.alamat,
            provinsi: body.provinsi,
            kabkota: body.kabkota,
            kecamatan: body.kecamatan,
            kelurahan: body.kelurahan,
            kode_pos: body.kode_pos,
            anakke: body.roleID == '3' ? body.anakke : null,
            jumlah_saudara: body.roleID == '3' ? body.jumlah_saudara : null,
            hobi: body.roleID == '3' ? body.hobi : null,
            cita_cita: body.roleID == '3' ? body.cita_cita : null,
            jenjang: body.roleID == '3' ? body.jenjang : null,
            status_sekolah: body.roleID == '3' ? body.status_sekolah : null,
            nama_sekolah: body.roleID == '3' ? body.nama_sekolah : null,
            npsn: body.roleID == '3' ? body.npsn : null,
            alamat_sekolah: body.roleID == '3' ? body.alamat_sekolah : null,
            kabkot_sekolah: body.roleID == '3' ? body.kabkot_sekolah : null,
            no_peserta_un: body.roleID == '3' ? body.no_peserta_un : null,
            no_skhun: body.roleID == '3' ? body.no_skhun : null,
            no_ijazah: body.roleID == '3' ? body.no_ijazah : null,
            nilai_un: body.roleID == '3' ? body.nilai_un : null,
            no_kk: body.roleID == '3' ? body.no_kk : null,
            nama_kk: body.roleID == '3' ? body.nama_kk : null,
            penghasilan: body.roleID == '3' ? body.penghasilan : null,
            nik_ayah: body.roleID == '3' ? body.nik_ayah : null,
            nama_ayah: body.roleID == '3' ? body.nama_ayah : null,
            tahun_ayah: body.roleID == '3' ? body.tahun_ayah : null,
            status_ayah: body.roleID == '3' ? body.status_ayah : null,
            pendidikan_ayah: body.roleID == '3' ? body.pendidikan_ayah : null,
            pekerjaan_ayah: body.roleID == '3' ? body.pekerjaan_ayah : null,
            telp_ayah: body.roleID == '3' ? body.telp_ayah : null,
            nik_ibu: body.roleID == '3' ? body.nik_ibu : null,
            nama_ibu: body.roleID == '3' ? body.nama_ibu : null,
            tahun_ibu: body.roleID == '3' ? body.tahun_ibu : null,
            status_ibu: body.roleID == '3' ? body.status_ibu : null,
            pendidikan_ibu: body.roleID == '3' ? body.pendidikan_ibu : null,
            pekerjaan_ibu: body.roleID == '3' ? body.pekerjaan_ibu : null,
            telp_ibu: body.roleID == '3' ? body.telp_ibu : null,
            nik_wali: body.roleID == '3' ? body.nik_wali : null,
            nama_wali: body.roleID == '3' ? body.nama_wali : null,
            tahun_wali: body.roleID == '3' ? body.tahun_wali : null,
            pendidikan_wali: body.roleID == '3' ? body.pendidikan_wali : null,
            pekerjaan_wali: body.roleID == '3' ? body.pekerjaan_wali : null,
            telp_wali: body.roleID == '3' ? body.telp_wali : null,
            status_tempat_tinggal: body.roleID == '3' ? body.status_tempat_tinggal : null,
            jarak_rumah: body.roleID == '3' ? body.jarak_rumah : null,
            transportasi: body.roleID == '3' ? body.transportasi : null,
            pendidikan_guru: body.roleID === '2' ? body.pendidikan_guru : null,
            jabatan_guru: body.roleID === '2' ? body.jabatan_guru : null,
            mengajar_bidang: body.roleID === '2' ? body.mengajar_bidang : null,
            mengajar_kelas: body.roleID === '2' ? body.mengajar_kelas : null,
            walikelas: body.roleID === '2' ? body.walikelas : null,
          }
          await MainModel.insert(kirimdata2, 'users_details');
          transporter = nodemailer.createTransport({
            host: 'smtp.gmail.com',
            port: 587,
            secure: false, // true for 465, false for other ports
            service: 'gmail',
            auth: {
              user: 'triyoga.ginanjar.p@gmail.com',
              pass: 'mdkqhwqbguuhggez' //Yoga17051993
            }
          });

          html = `<h1>Konfirmasi Pendataran Akun</h1>
          <ul>`;
            if(body.roleID !== 1) {`<li>Nomor Induk ${body.roleID === 2 ? 'Pegawai' : 'Siswa' } : ${body.nomor_induk}</li>`;}
          html += `<li>Nama Lengkap : ${body.name}</li>
            <li>Alamat Email : ${body.email}</li>
            <li>Kata Sandi : ${body.password}</li>
          </ul>
          Harap informasi ini jangan di hapus karena informasi ini penting adanya, dan klik tautan ini untuk mengonfirmasi pendaftaran Anda:<br>
          <a href="${process.env.BASE_URL}api/v1/moduleMain/verifikasi/${body.password}/1">konfirmasi akun</a><br>Jika Anda memiliki pertanyaan, silakan balas email ini`;
          
          mailOptions = {
            from: process.env.EMAIL,
            to: body.email,
            subject: 'Konfirmasi Pendaftaran Akun',
            // text: `Silahkan masukan kode verifikasi akun tersebut`
            html: html,
          };

          transporter.sendMail(mailOptions, (err, info) => {
            if (err) return response(res, { kode: 500, message: 'Gagal mengirim data ke alamat email anda, cek lagi email yang di daftarkan!.' }, 500);
          });
        break;
        case ('EDIT'):
          if(!readData.length){ return response(res, { kode: 404, message: 'Data tidak di temukan !' }, 404); }
          salt = await bcrypt.genSalt();
          hashPassword = await bcrypt.hash(body.kodeOTP, salt);
          const passbaru = body.password === readData[0].kodeOTP ? readData[0].password : hashPassword
          const kondisipassbaru = body.password === readData[0].kodeOTP ? `<b>Menggunakan kata sandi yang lama (${readData[0].kodeOTP})</b>` : body.password
          const kodeverifikasi = body.password === readData[0].kodeOTP ? readData[0].kodeOTP : body.kodeOTP
          kirimdata1 = {
            name: body.name,
            email: body.email,
            password: passbaru,
            activeAkun: '0',
            kodeOTP: body.password === readData[0].kodeOTP ? readData[0].kodeOTP : body.kodeOTP
          }
          await MainModel.update(kirimdata1, 'users', { id: body.id });
          kirimdata2 = {
						nomor_induk: body.roleID == '1' ? '' : body.nomor_induk,
						nik_siswa: body.roleID == '3' ? body.nik_siswa : null,
						tempat: body.tempat,
						tgl_lahir: body.tgl_lahir,
						jeniskelamin: body.jeniskelamin,
						agama: body.agama,
						telp: body.telp,
						alamat: body.alamat,
						provinsi: body.provinsi,
						kabkota: body.kabkota,
						kecamatan: body.kecamatan,
						kelurahan: body.kelurahan,
						kode_pos: body.kode_pos,
						anakke: body.roleID == '3' ? body.anakke : null,
						jumlah_saudara: body.roleID == '3' ? body.jumlah_saudara : null,
						hobi: body.roleID == '3' ? body.hobi : null,
						cita_cita: body.roleID == '3' ? body.cita_cita : null,
						jenjang: body.roleID == '3' ? body.jenjang : null,
						status_sekolah: body.roleID == '3' ? body.status_sekolah : null,
						nama_sekolah: body.roleID == '3' ? body.nama_sekolah : null,
						npsn: body.roleID == '3' ? body.npsn : null,
						alamat_sekolah: body.roleID == '3' ? body.alamat_sekolah : null,
						kabkot_sekolah: body.roleID == '3' ? body.kabkot_sekolah : null,
						no_peserta_un: body.roleID == '3' ? body.no_peserta_un : null,
						no_skhun: body.roleID == '3' ? body.no_skhun : null,
						no_ijazah: body.roleID == '3' ? body.no_ijazah : null,
						nilai_un: body.roleID == '3' ? body.nilai_un : null,
						no_kk: body.roleID == '3' ? body.no_kk : null,
						nama_kk: body.roleID == '3' ? body.nama_kk : null,
						penghasilan: body.roleID == '3' ? body.penghasilan : null,
						nik_ayah: body.roleID == '3' ? body.nik_ayah : null,
						nama_ayah: body.roleID == '3' ? body.nama_ayah : null,
						tahun_ayah: body.roleID == '3' ? body.tahun_ayah : null,
						status_ayah: body.roleID == '3' ? body.status_ayah : null,
						pendidikan_ayah: body.roleID == '3' ? body.pendidikan_ayah : null,
						pekerjaan_ayah: body.roleID == '3' ? body.pekerjaan_ayah : null,
						telp_ayah: body.roleID == '3' ? body.telp_ayah : null,
						nik_ibu: body.roleID == '3' ? body.nik_ibu : null,
						nama_ibu: body.roleID == '3' ? body.nama_ibu : null,
						tahun_ibu: body.roleID == '3' ? body.tahun_ibu : null,
						status_ibu: body.roleID == '3' ? body.status_ibu : null,
						pendidikan_ibu: body.roleID == '3' ? body.pendidikan_ibu : null,
						pekerjaan_ibu: body.roleID == '3' ? body.pekerjaan_ibu : null,
						telp_ibu: body.roleID == '3' ? body.telp_ibu : null,
						nik_wali: body.roleID == '3' ? body.nik_wali : null,
						nama_wali: body.roleID == '3' ? body.nama_wali : null,
						tahun_wali: body.roleID == '3' ? body.tahun_wali : null,
						pendidikan_wali: body.roleID == '3' ? body.pendidikan_wali : null,
						pekerjaan_wali: body.roleID == '3' ? body.pekerjaan_wali : null,
						telp_wali: body.roleID == '3' ? body.telp_wali : null,
						status_tempat_tinggal: body.roleID == '3' ? body.status_tempat_tinggal : null,
						jarak_rumah: body.roleID == '3' ? body.jarak_rumah : null,
						transportasi: body.roleID == '3' ? body.transportasi : null,
						pendidikan_guru: body.roleID === '2' ? body.pendidikan_guru : null,
						jabatan_guru: body.roleID === '2' ? body.jabatan_guru : null,
						mengajar_bidang: body.roleID === '2' ? body.mengajar_bidang : null,
						mengajar_kelas: body.roleID === '2' ? body.mengajar_kelas : null,
						walikelas: body.roleID === '2' ? body.walikelas : null,
					}
          await MainModel.update(kirimdata2, 'users_details', { id_profile: body.id });
          transporter = nodemailer.createTransport({
            host: 'smtp.gmail.com',
            port: 587,
            secure: false, // true for 465, false for other ports
            service: 'gmail',
            auth: {
              user: 'triyoga.ginanjar.p@gmail.com',
              pass: 'mdkqhwqbguuhggez' //Yoga17051993
            }
          });

          html = `<h1>Konfirmasi Perubahan Data Akun</h1>
          <ul>`;
            if(body.roleID !== 1) {`<li>Nomor Induk ${body.roleID === 2 ? 'Pegawai' : 'Siswa' } : ${body.nomor_induk}</li>`;}
          html += `<li>Nama Lengkap : ${body.name}</li>
            <li>Alamat Email : ${body.email}</li>
            <li>Kata Sandi : ${kondisipassbaru}</li>
          </ul>
          Harap informasi ini jangan di hapus karena informasi ini penting adanya, dan klik tautan ini untuk mengonfirmasi pendaftaran Anda:<br>
          <a href="${process.env.BASE_URL}api/v1/moduleMain/verifikasi/${kodeverifikasi}/1">konfirmasi akun</a><br>Jika Anda memiliki pertanyaan, silakan balas email ini`;
          
          mailOptions = {
            from: process.env.EMAIL,
            to: body.email,
            subject: 'Konfirmasi Perubahan Data Akun',
            // text: `Silahkan masukan kode verifikasi akun tersebut`
            html: html,
          };

          transporter.sendMail(mailOptions, (err, info) => {
            if (err) return response(res, { kode: 500, message: 'Gagal mengirim data ke alamat email anda, cek lagi email yang di daftarkan!.' }, 500);
          });
        break;
        default:
          logger.error(`[ERROR ACCESS] --- Gagal memproses data !`);
          return response(res, { kode: 404, message: 'Gagal memproses data !' }, 404);
      }
      let kode = 200;
      let message = 'Berhasil';
      logger.info(`[${(body.jenis == 'ADD')?'INSERT':'UPDATE'} DATA] --- ${(body.jenis == 'ADD')?'insert':'update'} data users berhasil !`);
      response(res, { kode, message }, 200);
    }catch(err){
      console.log(err)
      logger.error(`[ERROR ACCESS] --- Gagal memproses data !`);
      return response(res, { kode: 404, message: 'Gagal memproses data !' }, 404);
    }
  }
  
  updateProfileData = async (req, res, next) => {
    try{
      let body = { ...req.body };
      let attributes = []
      let include = {}
      let search = { or: {}, and: { id: body.id } }

      let data = await this.readDataBy(attributes, search, include, 'viewusers.id', 'viewusers');
      if(data.length){
        if(body.ubah === 'nama'){
          const kirimData = {
              name: body.name
          }
          await MainModel.update(kirimData, 'users', { id: body.id });
        }else if(body.ubah === 'katasandi'){
          if (body.passwordlama == '' || body.passwordlama == null) { return response(res, { kode: 404, message: 'Kata Sandi Lama tidak boleh kosong' }, 404); }
          const match = await bcrypt.compare(body.passwordlama, data[0].password);
          if(!match) return response(res, { kode: 404, message: 'Kata Sandi Lama salah !' }, 404);
          if (body.passwordbaru == '' || body.passwordbaru == null) { return response(res, { kode: 404, message: 'Kata Sandi Baru tidak boleh kosong' }, 404); }
          else if (body.confPasswordbaru == '' || body.confPasswordbaru == null) { return response(res, { kode: 404, message: 'Konfirmasi Kata Sandi Baru tidak boleh kosong' }, 404); }
          else if(body.passwordbaru !== body.confPasswordbaru) return response(res, { kode: 404, message: 'Kata Sandi dan Konfirmasi Kata Sandi tidak cocok !' }, 404);
          const salt = await bcrypt.genSalt();
          const hashPassword = await bcrypt.hash(body.passwordbaru, salt);

          const userID = data[0].id;
          const name = data[0].name;
          const email = data[0].email;
          const codeLog = data[0].codeLog;
          const accessToken = jwt.sign({userID, name, email}, process.env.ACCESS_TOKEN_SECRET, {
                  expiresIn: '12h'
          });
          const refreshToken = jwt.sign({userID, name, email}, process.env.REFRESH_TOKEN_SECRET, {
                  expiresIn: '1d'
          });

          const kirimData = {
              password: hashPassword,
              refresh_token: refreshToken,
              codeLog: codeLog,
              gambarGmail: null,
              kodeOTP: body.passwordbaru
          }
          await MainModel.update(kirimData, 'users', { id: body.id });
        }else if(body.ubah === 'datapribadi'){
          const kirimData = {
            email: body.email
          }
          await MainModel.update(kirimData, 'users', { id: body.id });

          const kirimData2 = {
            telp: body.telp,
            alamat: body.alamat,
            provinsi: body.provinsi,
            kabkota: body.kabkota,
            kecamatan: body.kecamatan,
            kelurahan: body.kelurahan,
            kode_pos: body.kode_pos,
          }
          await MainModel.update(kirimData2, 'users_details', { id_profile: body.id });
        }else if(body.ubah === 'loginGmail'){
          const kirimData = {
            gambarGmail: body.gambarGmail,
            codeLog: '2'
          }
          await MainModel.update(kirimData, 'users', { email: body.email });
        }

        let kode = 200;
        let message = 'Berhasil';
        logger.info(`[UPDATED DATA] --- update data users berhasil !`);
        response(res, { kode, message }, 200);
      }else{
        return response(res, { kode: 404, message: 'Data tidak di temukan !' }, 404);
      }
    }catch(err){
      logger.error(`[ERROR ACCESS] --- Gagal memproses data !`);
      return response(res, { kode: 404, message: 'Gagal memproses data !' }, 404);
    }
  }

  deleteData = async (req, res, next) => {
    try{
      await MainModel.delete({ id: req.params.id }, 'users');
      let kode = 200;
      let message = 'Berhasil';
      logger.info(`[DELETED DATA] --- delete data users berhasil !`);
      response(res, { kode, message }, 200);
    }catch(err){
      logger.error(`[ERROR ACCESS] --- Gagal memproses data !`);
      return response(res, { kode: 404, message: 'Gagal memproses data !' }, 404);
    }
  }

  verifikasi = async (req, res, next) => {
    try{
      let params = { ...req.params };  

      let search = {
        or: {},
        and: {
          kodeOTP: params.kode
        },
      }
      let readData = await this.readDataBy(search, 'viewusers');
      if(!readData.length){ return response(res, { kode: 404, message: 'Data tidak di temukan !' }, 404); }
      const kirimData = {
				gambarGmail : null,
				refresh_token : null,
				codeLog : '0',
				activeAkun : params.activeAkun
			}
      await MainModel.update(kirimData, 'users', { id: readData[0].id });
      res.send("<script>window.close();</script > ")
    }catch(err){
      logger.error(`[ERROR ACCESS] --- Gagal memproses data !`);
      return response(res, { kode: 404, message: 'Gagal memproses data !' }, 404);
    }
  }

  updateFile = async (req, res, next) => {
    try{
      let namaFile = req.files[0].filename;
      let body = { ...req.body, namaFile };
      let kirimData = {
        gambar: body.namaFile
      }
      await MainModel.update(kirimData, 'users', { id: body.id });
      let kode = 200;
        let message = 'Berhasil';
        logger.info(`[UPDATED DATA] --- update data users berhasil !`);
        response(res, { kode, message }, 200);
    }catch(err){
      logger.error(`[ERROR ACCESS] --- Gagal memproses data !`);
      return response(res, { kode: 404, message: 'Gagal memproses data !' }, 404);
    }
  }

  updateBerkas = async (req, res, next) => {
    try{
      let namaFile = req.files[0].filename;
      let body = { ...req.body, namaFile };
      let attributes = []
      let include = {}
      let search = {
        or: {},
        and: {
          id_profile: body.id
        },
      }

      let kirimData
      if(body.namaBerkas === 'ijazah'){
        kirimData = {fc_ijazah: body.namaFile}
      }else if(body.namaBerkas === 'kk'){
        kirimData = {fc_kk: body.namaFile}
      }else if(body.namaBerkas === 'ktp'){
        kirimData = {fc_ktp_ortu: body.namaFile}
      }else if(body.namaBerkas === 'aktalahir'){
        kirimData = {fc_akta_lahir: body.namaFile}
      }else if(body.namaBerkas === 'skl'){
        kirimData = {fc_skl: body.namaFile}
      }
      await MainModel.update(kirimData, 'users_details', { id_profile: body.id });
      let readData = await this.readDataBy(attributes, search, include, 'viewusers.id', 'viewusers');
      let kode = 200;
      let message = 'Berhasil';
      logger.info(`[UPDATED DATA] --- update data users berhasil !`);
      response(res, { kode, message, result: readData[0] }, 200);
    }catch(err){
      logger.error(`[ERROR ACCESS] --- Gagal memproses data !`);
      return response(res, { kode: 404, message: 'Gagal memproses data !' }, 404);
    }
  }

  readDataProvinsi = async (req, res, next) => {
    try{
      let prototype = {
        field: 'provinsi',
        search: {
          kodeLength: '2'
        }
      }
      let data = await MainModel.findWilayah(prototype);
      let kode = 200;
      let message = 'Berhasil';
      response(res, { kode, message, result: data }, 200);
    }catch(err){
      console.log(err)
      logger.error(`[ERROR ACCESS] --- Gagal memproses data !`);
      return response(res, { kode: 404, message: 'Gagal memproses data !' }, 404);
    }
  }

  readDataKabKotaOnly = async (req, res, next) => {
    try{
      let prototype = {
        field: 'kabkotaOnly',
        search: {
          kodeLength: '5'
        }
      }
      let data = await MainModel.findWilayah(prototype);
      let kode = 200;
      let message = 'Berhasil';
      response(res, { kode, message, result: data }, 200);
    }catch(err){
      logger.error(`[ERROR ACCESS] --- Gagal memproses data !`);
      return response(res, { kode: 404, message: 'Gagal memproses data !' }, 404);
    }
  }

  readDataKabKota = async (req, res, next) => {
    try{
      let params = req.params.id
      let jmlString = params.length
      let whereChar = (jmlString==2?5:(jmlString==5?8:13))
      let search = { jmlString: jmlString, kodeWilayah: params, kodeLength: whereChar };
      let prototype = {
        field: 'kabkota',
        search: search
      }
      let data = await MainModel.findWilayah(prototype);
      let kode = 200;
      let message = 'Berhasil';
      response(res, { kode, message, result: data }, 200);
    }catch(err){
      logger.error(`[ERROR ACCESS] --- Gagal memproses data !`);
      return response(res, { kode: 404, message: 'Gagal memproses data !' }, 404);
    }
  }

  readDataKecamatan = async (req, res, next) => {
    try{
      let params = req.params.id
      let jmlString = params.length
      let whereChar = (jmlString==2?5:(jmlString==5?8:13))
      let search = { jmlString: jmlString, kodeWilayah: params, kodeLength: whereChar };
      let prototype = {
        field: 'kabkota',
        search: search
      }
      let data = await MainModel.findWilayah(prototype);
      let kode = 200;
      let message = 'Berhasil';
      response(res, { kode, message, result: data }, 200);
    }catch(err){
      logger.error(`[ERROR ACCESS] --- Gagal memproses data !`);
      return response(res, { kode: 404, message: 'Gagal memproses data !' }, 404);
    }
  }

  readDataKelDesa = async (req, res, next) => {
    try{
      let params = req.params.id
      let jmlString = params.length
      let whereChar = (jmlString==2?5:(jmlString==5?8:13))
      let search = { jmlString: jmlString, kodeWilayah: params, kodeLength: whereChar };
      let prototype = {
        field: 'kabkota',
        search: search
      }
      let data = await MainModel.findWilayah(prototype);
      let kode = 200;
      let message = 'Berhasil';
      response(res, { kode, message, result: data }, 200);
    }catch(err){
      logger.error(`[ERROR ACCESS] --- Gagal memproses data !`);
      return response(res, { kode: 404, message: 'Gagal memproses data !' }, 404);
    }
  }

  getkelas = async (req, res, next) => {
    try{
      let query = { ...req.query }
      let search = query.kelas == 'ALL' ? 
        {
          or: {},
          and: {
            activeKelas: 1
          },
        }
      :
        {
          or: {},
          and: {
            kelas: query.kelas,
            activeKelas: 1
          },
        }
      let data = await MainModel.findKelas(query, search)
      let kode = 200;
      let message = 'Berhasil';
      response(res, { kode, message, result: data }, 200);
    }catch(err){
      logger.error(`[ERROR ACCESS] --- Gagal memproses data !`);
      return response(res, { kode: 404, message: 'Gagal memproses data !' }, 404);
    }
  }

  ambilKelas = async (req, res, next) => {
    try{
      let body = { ...req.body };
      let attributes = []
      let include = {}
      let search = {
        or: {},
        and: {
          id: body.id
        },
      }
      let searchNilai = {
        or: {},
        and: {
          id_profile: body.id
        },
      }
      let data = await this.readDataBy(attributes, search, include, 'viewusers.id', 'viewusers')
      if(!data.length){ return response(res, { kode: 404, message: 'Data tidak di temukan !' }, 404); }
      await MainModel.update({ kelas: body.kelas }, 'users_details', { id_profile: body.id });
      let cekdataNilai = await this.readDataBy(attributes, searchNilai, include, 'nilai.id_nilai', 'nilai')
      if(!cekdataNilai.length) {
        const mapel = ['Alquran Hadits', 'Aqidah Akhlak', 'Bahasa Arab', 'Bahasa Indonesia', 'Bahasa Inggris',
                'Bahasa Sunda', 'Fiqih', 'IPA Terpadu', 'IPS Terpadu', 'Matematika', 'Penjasorkes', 'PKN',
                'Prakarya', 'Seni Budaya', 'SKI']
        for(let i=0;i<mapel.length;i++){
          const simpanData = {
            id_profile: body.id,
            mapel: mapel[i]
          }
          await MainModel.insert(simpanData, 'nilai');
        }
      }
      let readData = await this.readDataBy(attributes, search, include, 'viewusers.id', 'viewusers')
      let hasilData = {
        data: readData[0],
        nilai: cekdataNilai
      }
      let kode = 200;
      let message = 'Berhasil';
      response(res, { kode, message, result: hasilData }, 200);
    }catch(err){
      console.log(err)
      logger.error(`[ERROR ACCESS] --- Gagal memproses data !`);
      return response(res, { kode: 404, message: 'Gagal memproses data !' }, 404);
    }
  }

  kelasSiswa = async (req, res, next) => {
    try{
      let kelas = req.params.kelas
      let attributes = []
      let include = {}
      let search = {
        or: {},
        and: {
          activeAkun: 1,
          mutationAkun: 0,
          kelas: kelas
        },
      }
      let data = await this.readDataBy(attributes, search, include, 'viewusers.id', 'viewusers');
      let kode = 200;
      let message = 'Berhasil';
      response(res, { kode, message, result: data }, 200);
    }catch(err){
      console.log(err)
      logger.error(`[ERROR ACCESS] --- Gagal memproses data !`);
      return response(res, { kode: 404, message: 'Gagal memproses data !' }, 404);
    }
  }

  penilaianSiswa = async (req, res, next) => {
    try{
      let { mapel, kelas } = req.query
      let attributesOne = []
      let includeOne = {
        attributes: ['mapel', 'kelas_mengajar'],
        joinTable: 'jadwal_mengajar',
        joinOn: 'jadwal_mengajar.kelas_mengajar'
      }
      let searchOne = {
        or: {},
        and: {
          activeAkun: 1,
          mutationAkun: 0,
          mapel: mapel,
          kelas_mengajar: kelas,
        }
      }
      let data = await this.readDataBy(attributesOne, searchOne, includeOne, 'viewusers.kelas', 'viewusers');
      let dataTampung = await Promise.all(data.map( async (value) => {
        let attributesTwo = []
        let includeTwo = {}
        let searchTwo = {
          or: {},
          and: {
            id_profile: value.id,
            mapel: mapel,
          }
        }
        let dataNilai = await this.readDataBy(attributesTwo, searchTwo, includeTwo, 'nilai.id_nilai', 'nilai');
        let tampung = await Promise.all(dataNilai.map((value2) => {
          let tampungData = []
          if(value.id == value2.id_profile){
            // console.log(value, value2)
            let objectBaru = Object.assign(value, value2);
            tampungData.push(objectBaru)
          }
          return tampungData[0]
        }))
        return tampung[0];
      }))
      let kode = 200;
      let message = 'Berhasil';
      response(res, { kode, message, result: dataTampung }, 200);
    }catch(err){
      console.log(err)
      logger.error(`[ERROR ACCESS] --- Gagal memproses data !`);
      return response(res, { kode: 404, message: 'Gagal memproses data !' }, 404);
    }
  }

  ubahPenilaian = async (req, res, next) => {
    try{
      let body = { ...req.body }
      let dataNilai = body.ubahNilai
      if(!body.triggerUbah) return response(res, { kode: '404', message: 'Anda belum memilih nilai yang ingin di ubah' }, 404);
      for(let i=0;i<dataNilai.length;i++){
        let simpanData
        if(body.triggerUbah == 'Tugas 1'){
          simpanData = { n_tugas1: dataNilai[i].nilai ? dataNilai[i].nilai : null }
        }else if(body.triggerUbah == 'Tugas 2'){
          simpanData = { n_tugas2: dataNilai[i].nilai ? dataNilai[i].nilai : null }
        }else if(body.triggerUbah == 'Tugas 3'){
          simpanData = { n_tugas3: dataNilai[i].nilai ? dataNilai[i].nilai : null }
        }else if(body.triggerUbah == 'Tugas 4'){
          simpanData = { n_tugas4: dataNilai[i].nilai ? dataNilai[i].nilai : null }
        }else if(body.triggerUbah == 'Tugas 5'){
          simpanData = { n_tugas5: dataNilai[i].nilai ? dataNilai[i].nilai : null }
        }else if(body.triggerUbah == 'Tugas 6'){
          simpanData = { n_tugas6: dataNilai[i].nilai ? dataNilai[i].nilai : null }
        }else if(body.triggerUbah == 'Tugas 7'){
          simpanData = { n_tugas7: dataNilai[i].nilai ? dataNilai[i].nilai : null }
        }else if(body.triggerUbah == 'Tugas 8'){
          simpanData = { n_tugas8: dataNilai[i].nilai ? dataNilai[i].nilai : null }
        }else if(body.triggerUbah == 'Tugas 9'){
          simpanData = { n_tugas9: dataNilai[i].nilai ? dataNilai[i].nilai : null }
        }else if(body.triggerUbah == 'Tugas 10'){
          simpanData = { n_tugas10: dataNilai[i].nilai ? dataNilai[i].nilai : null }
        }else if(body.triggerUbah == 'UTS'){
          simpanData = { n_uts: dataNilai[i].nilai ? dataNilai[i].nilai : null }
        }else if(body.triggerUbah == 'UAS'){
          simpanData = { n_uas: dataNilai[i].nilai ? dataNilai[i].nilai : null }
        }
        await MainModel.update(simpanData, 'nilai', { id_profile: dataNilai[i].id_profile, mapel: body.mapel });
      }

      let kode = 200;
      let message = 'Berhasil';
      response(res, { kode, message }, 200);
    }catch(err){
      console.log(err)
      logger.error(`[ERROR ACCESS] --- Gagal memproses data !`);
      return response(res, { kode: 404, message: 'Gagal memproses data !' }, 404);
    }
  }

  getjadwalNgajar = async (req, res, next) => {
    try{
      let param = { ...req.params }
      let attributes = []
      let include = {
        attributes: ['name', 'nomor_induk'],
        joinTable: 'viewusers',
        joinOn: 'viewusers.id'
      }
      let search = {
        or: {},
        and: {
          id: param.id,
        }
      }
      let data = await this.readDataBy(attributes, search, include, 'jadwal_mengajar.id_profile', 'jadwal_mengajar');
      let kode = 200;
      let message = 'Berhasil';
      response(res, { kode, message, result: data }, 200);
    }catch(err){
      console.log(err)
      logger.error(`[ERROR ACCESS] --- Gagal memproses data !`);
      return response(res, { kode: 404, message: 'Gagal memproses data !' }, 404);
    }
  }

  createJadwalNgajar = async (req, res, next) => {
    try{
      let body = { ...req.body }
      let attributes = []
      let include = {}
      let search = {
        or: {},
        and: {
          id_profile: body.id,
          mapel: body.mapel,
          kelas_mengajar: body.kelas,
          status: 1
        },
      }
      let data = await this.readDataBy(attributes, search, include, 'jadwal_mengajar.id_jadwal', 'jadwal_mengajar');
      if(data.length) return response(res, { kode: 404, message: 'Jadwal Mengajar sudah ada !' }, 404);
      const kirimData = {
				id_profile: body.id,
				mapel: body.mapel,
				kelas_mengajar: body.kelas,
				status: '1'
			}
      await MainModel.insert(kirimData, 'jadwal_mengajar');
      let kode = 200;
      let message = 'Berhasil';
      response(res, { kode, message }, 200);
    }catch(err){
      console.log(err)
      logger.error(`[ERROR ACCESS] --- Gagal memproses data !`);
      return response(res, { kode: 404, message: 'Gagal memproses data !' }, 404);
    }
  }

  deleteJadwalNgajar = async (req, res, next) => {
    try{
      await MainModel.delete({ id_jadwal: req.params.id_jadwal }, 'jadwal_mengajar');
      let kode = 200;
      let message = 'Berhasil';
      response(res, { kode, message }, 200);
    }catch(err){
      logger.error(`[ERROR ACCESS] --- Gagal memproses data !`);
      return response(res, { kode: 404, message: 'Gagal memproses data !' }, 404);
    }
  }

  updateStatus = async (req, res, next) => {
    try{
      let body = { ...req.body }
      let attributes = []
      let include = {}
      let where = body.table == 'users' ? { id: body.id } : { id_jadwal: body.id }
      let search = {
        or: {},
        and: where,
      }
      let pesan = '', kirimData
      if(body.table === 'users') {
				switch(body.jenis) {
					case 'activeAkun' :
							kirimData = {
								activeAkun: body.activeAkun
							}
							pesan = body.activeAkun === '0' ? 'activeAkunNot' : 'activeAkun'
						break;
					case 'validasiAkun' :
							kirimData = {
								validasiAkun: body.validasiAkun
							}
							pesan = body.validasiAkun === '0' ? 'validasiAkunNot' : 'validasiAkun'
						break;
					case 'mutationAkun' :
							kirimData = {
								activeAkun: body.activeAkun,
								mutationAkun: body.mutationAkun
							}
							pesan = body.mutationAkun === '0' ? 'mutationAkunNot' : 'mutationAkun'
						break;
					default:
						console.log('Error')
				}
			}else{
				kirimData = {
					status: body.activeStatus
				}
				pesan = body.activeStatus === '0' ? 'statusJadwalNot' : 'statusJadwal'
			}

      const nomeklatur_pesan = {
        activeAkunNot: 'Berhasil mengubah aktif Akun menjadi tidak aktif',
        activeAkun: 'Berhasil mengubah aktif Akun menjadi aktif',
        validasiAkunNot: 'Berhasil mengubah data Akun menjadi tidak tervalidasi',
        validasiAkun: 'Berhasil mengubah data Akun menjadi tervalidasi',
        mutationAkunNot: 'Berhasil mengubah data Akun menjadi tidak di mutasi',
        mutationAkun: 'Berhasil mengubah data Akun menjadi di mutasi',
        statusJadwalNot: 'Berhasil mengubah aktif Jadwal Mengajar menjadi tidak aktif',
        statusJadwal: 'Berhasil mengubah aktif Jadwal Mengajar menjadi aktif',
      }[pesan];

      let data = await this.readDataBy(attributes, search, include, `${ body.table == 'users' ? `${body.table}.id` : `${body.table}.id_jadwal` }`, body.table);
      if(!data.length) return response(res, { kode: 404, message: 'Data tidak di temukan !' }, 404);
      await MainModel.update(kirimData, body.table, where);
      let kode = 200;
      let message = nomeklatur_pesan;
      response(res, { kode, message }, 200);
    }catch(err){
      console.log(err)
      logger.error(`[ERROR ACCESS] --- Gagal memproses data !`);
      return response(res, { kode: 404, message: 'Gagal memproses data !' }, 404);
    }
  }
  
  detailUserPDF = async (req, res, next) => {
    try{
      let id = req.params.id;
      let attributes = []
      let include = {}
      let search = {
        or: {},
        and: {
          id: id
        },
      }

      let data = await this.readDataBy(attributes, search, include, 'viewusers.id', 'viewusers');
      const optionsAgama = [
        { value: 'Islam', label: 'Islam' },
        { value: 'Katolik', label: 'Katolik' },
        { value: 'Protestan', label: 'Protestan' },
        { value: 'Hindu', label: 'Hindu' },
        { value: 'Budha', label: 'Budha' },
      ]
    
      const optionsHobi = [
        { value: '1', label: 'Olahraga' },
        { value: '2', label: 'Kesenian' },
        { value: '3', label: 'Membaca' },
        { value: '4', label: 'Menulis' },
        { value: '5', label: 'Traveling' },
        { value: '6', label: 'Lainnya' },
      ]
      
      const optionsCitaCita = [
        { value: '1', label: 'PNS' },
        { value: '2', label: 'TNI/PORLI' },
        { value: '3', label: 'Guru/Dosen' },
        { value: '4', label: 'Dokter' },
        { value: '5', label: 'Politikus' },
        { value: '6', label: 'Wiraswasta' },
        { value: '7', label: 'Pekerja Seni/Lukis/Artis/Sejenis' },
        { value: '8', label: 'Lainnya' },
      ]
      
      const optionsJenjang = [
        { value: '1', label: 'MI' },
        { value: '2', label: 'SD' },
        { value: '3', label: 'SD Terbuka' },
        { value: '4', label: 'SLB-MI' },
        { value: '5', label: 'Paket A' },
        { value: '6', label: 'Salafiyah Ula' },
        { value: '7', label: 'MU`adalah MI' },
        { value: '8', label: 'SLB-SD' },
        { value: '9', label: 'Lainnya' },
      ]
    
      const optionsStatusSekolah = [
        { value: '1', label: 'Negeri' },
        { value: '2', label: 'Swasta' },
      ]
    
      const optionsStatusOrtu = [
        { value: '1', label: 'Masih Hidup' },
        { value: '2', label: 'Sudah Mati' },
        { value: '3', label: 'Tidak Diketahui' },
      ]
      
      const optionsPendidikan = [
        { value: '0', label: 'Tidak Berpendidikan Formal' },
        { value: '1', label: 'SD/Sederajat' },
        { value: '2', label: 'SMP/Sederajat' },
        { value: '3', label: 'SMA/Sederajat' },
        { value: '4', label: 'D1' },
        { value: '5', label: 'D2' },
        { value: '6', label: 'D3' },
        { value: '7', label: 'S1' },
        { value: '8', label: 'S2' },
        { value: '9', label: '>S2' },
      ]
    
      const optionsPekerjaan = [
        { value: '1', label: 'Tidak Bekerja' },
        { value: '2', label: 'Pensiunan/Almarhum' },
        { value: '3', label: 'PNS (selain Guru/Dosen/Dokter/Bidan/Perawat)' },
        { value: '4', label: 'TNI/Polisi' },
        { value: '5', label: 'Guru/Dosen' },
        { value: '6', label: 'Pegawai Swasta' },
        { value: '7', label: 'Pengusaha/Wiraswasta' },
        { value: '8', label: 'Pengacara/Hakim/Jaksa/Notaris' },
        { value: '9', label: 'Seniman/Pelukis/Artis/Sejenis' },
        { value: '10', label: 'Dokter/Bidan/Perawat' },
        { value: '11', label: 'Pilot/Pramugari' },
        { value: '12', label: 'Pedagang' },
        { value: '13', label: 'Petani/Peternak' },
        { value: '14', label: 'Nelayan' },
        { value: '15', label: 'Buruh (Tani/Pabrik/Bangunan)' },
        { value: '16', label: 'Sopir/Masinis/Kondektur' },
        { value: '17', label: 'Politikus' },
        { value: '18', label: 'Lainnya' },
      ]
    
      const optionsStatusTempatTinggal = [
        { value: '1', label: 'Milik' },
        { value: '2', label: 'Rumah Orangtua' },
        { value: '3', label: 'Rumah Saudara/Kerabat' },
        { value: '4', label: 'Rumah Dinas' },
      ]
    
      const optionsJarakRumah = [
        { value: '1', label: '< 1 Km' },
        { value: '2', label: '1 - 3 Km' },
        { value: '3', label: '3 - 5 Km' },
        { value: '4', label: '5 - 10 Km' },
        { value: '5', label: '> 10 Km' },
      ]
    
      const optionsAlatTransportasi = [
        { value: '1', label: 'Jalan Kaki' },
        { value: '2', label: 'Sepeda' },
        { value: '3', label: 'Sepeda Motor' },
        { value: '4', label: 'Mobil Pribadi' },
        { value: '5', label: 'Antar Jemput Sekolah' },
        { value: '6', label: 'Angkutan Umum' },
        { value: '7', label: 'Perahu/Sampan' },
        { value: '8', label: 'Lainnya' },
      ]
    
      const optionsPenghasilan = [
        { value: '1', label: '<= Rp 500.000' },
        { value: '2', label: 'Rp 500.001 - Rp 1.000.000' },
        { value: '3', label: 'Rp 1.000.001 - Rp 2.000.000' },
        { value: '4', label: 'Rp 2.000.001 - Rp 3.000.000' },
        { value: '5', label: 'Rp 3.000.001 - Rp 5.000.000' },
        { value: '6', label: '> Rp 5.000.000' },
      ]
  
      const agama = optionsAgama.find(dataagama => dataagama.value === String(data[0].agama));
      const citacita = optionsCitaCita.find(datacitacita => datacitacita.value === String(data[0].cita_cita));
      const hobi = optionsHobi.find(datahobi => datahobi.value === String(data[0].hobi));
      const jenjangsekolah = optionsJenjang.find(datajenjang => datajenjang.value === String(data[0].jenjang));
      const statussekolah = optionsStatusSekolah.find(datastasek => datastasek.value === String(data[0].status_sekolah));
      const penghasilan = optionsPenghasilan.find(datapenghasilan => datapenghasilan.value === String(data[0].penghasilan));
      const statusayah = optionsStatusOrtu.find(datastatusayah => datastatusayah.value === String(data[0].status_ayah));
      const statusibu = optionsStatusOrtu.find(datastatusibu => datastatusibu.value === String(data[0].status_ibu));
      const pendidikanayah = optionsPendidikan.find(datapendidikanayah => datapendidikanayah.value === String(data[0].pendidikan_ayah));
      const pendidikanibu = optionsPendidikan.find(datapendidikanibu => datapendidikanibu.value === String(data[0].pendidikan_ibu));
      const pendidikanwali = optionsPendidikan.find(datapendidikanwali => datapendidikanwali.value === String(data[0].pendidikan_wali));
      const pekerjaanayah = optionsPekerjaan.find(datapekerjaanayah => datapekerjaanayah.value === String(data[0].pekerjaan_ayah));
      const pekerjaanibu = optionsPekerjaan.find(datapekerjaanibu => datapekerjaanibu.value === String(data[0].pekerjaan_ibu));
      const pekerjaanwali = optionsPekerjaan.find(datapekerjaanwali => datapekerjaanwali.value === String(data[0].pekerjaan_wali));
      const statustempattinggal = optionsStatusTempatTinggal.find(datastatustempattinggal => datastatustempattinggal.value === String(data[0].status_tempat_tinggal));
      const jarakrumah = optionsJarakRumah.find(datajarakrumah => datajarakrumah.value === String(data[0].jarak_rumah));
      const transportasi = optionsAlatTransportasi.find(datatransportasi => datatransportasi.value === String(data[0].transportasi));
      const hasil = {
        ...data[0], 
        linkGatsa: `${process.env.BASE_URL}bahan/gatsa.png`,
        name: UpperFirstLetter(data[0].name),
        tempat: UpperFirstLetter(data[0].tempat),
        alamat: UpperFirstLetter(data[0].alamat),
        nama_sekolah: UpperFirstLetter(data[0].nama_sekolah),
        nama_kk: UpperFirstLetter(data[0].nama_kk),
        nama_ayah: UpperFirstLetter(data[0].nama_ayah),
        nama_ibu: UpperFirstLetter(data[0].nama_ibu),
        nama_wali: data[0].nama_wali ? UpperFirstLetter(data[0].nama_wali) : null,
        nama_provinsi: UpperFirstLetter(data[0].nama_provinsi), 
        nama_kabkota: UpperFirstLetter(data[0].nama_kabkota),
        nama_kabkot_sekolah: UpperFirstLetter(data[0].nama_kabkot_sekolah),
        nama_kecamatan: UpperFirstLetter(data[0].nama_kecamatan),
        nama_kelurahan: UpperFirstLetter(data[0].nama_kelurahan),
        tgl_lahir: dateconvert(data[0].tgl_lahir),
        agama: agama ? agama.label : '-',
        cita_cita: citacita ? citacita.label : '-',
        hobi: hobi ? hobi.label : '-',
        jenjang: jenjangsekolah ? jenjangsekolah.label : '-',
        jenjang: jenjangsekolah ? jenjangsekolah.label : '-',
        status_sekolah: statussekolah ? statussekolah.label : '-',
        penghasilan: penghasilan ? penghasilan.label : '-',
        status_ayah: statusayah ? statusayah.label : '-',
        status_ibu: statusibu ? statusibu.label : '-',
        pendidikan_ayah: pendidikanayah ? pendidikanayah.label : '-',
        pendidikan_ibu: pendidikanibu ? pendidikanibu.label : '-',
        pendidikan_wali: pendidikanwali ? pendidikanwali.label : '-',
        pekerjaan_ayah: pekerjaanayah ? pekerjaanayah.label : '-',
        pekerjaan_ibu: pekerjaanibu ? pekerjaanibu.label : '-',
        pekerjaan_wali: pekerjaanwali ? pekerjaanwali.label : '-',
        status_tempat_tinggal: statustempattinggal ? statustempattinggal.label : '-',
        jarak_rumah: jarakrumah ? jarakrumah.label : '-',
        transportasi: transportasi ? transportasi.label : '-',
      }
      // console.log(hasil)
      ejs.renderFile(path.join(__dirname, "../../src/views/viewSiswa.ejs"),{dataSiswa: hasil}, (err, data) => {
        if (err) {
          console.log(err)
        } else {
          // console.log(data)
          let options = {
            format: "A4",
            orientation: "portrait",
            quality: "10000",
            border: {
              top: "1.8cm",            // default is 0, units: mm, cm, in, px
              right: "2cm",
              bottom: "1.5cm",
              left: "2cm"
            },
            // header: {
            // 	height: "12mm",
            // },
            // footer: {
            // 	height: "15mm",
            // },
            httpHeaders: {
              "Content-type": "application/pdf",
            },
            type: "pdf",
  
          };
          pdf.create(data, options).toStream(function(err, stream){
            stream.pipe(res);
          });
        }
      });
    }catch(err){
      console.log(err)
      logger.error(`[ERROR ACCESS] --- Gagal memproses data !`);
      return response(res, { kode: 404, message: 'Gagal memproses data !' }, 404);
    }
  }  

  downloadexcel = async (req, res, next) => {
    try{
      let roleid = req.params.roleid;
      let workbook = new excel.Workbook();
      workbook.creator = 'Triyoga Ginanjar Pamungkas';
      workbook.created = new Date();
      if(roleid === '3'){
        let worksheet = workbook.addWorksheet("Data Siswa");
        let worksheetAgama = workbook.addWorksheet("Agama");
        let worksheetHobi = workbook.addWorksheet("Hobi");
        let worksheetCitaCita = workbook.addWorksheet("Cita - Cita");
        let worksheetJenjangSekolah = workbook.addWorksheet("Jenjang Sekolah");
        let worksheetStatusSekolah = workbook.addWorksheet("Status Sekolah");
        let worksheetStatusOrangTua = workbook.addWorksheet("Status Orang Tua");
        let worksheetPendidikan = workbook.addWorksheet("Pendidikan");
        let worksheetPekerjaan = workbook.addWorksheet("Pekerjaan");
        let worksheetStatusTempatTinggal = workbook.addWorksheet("Status Tempat Tinggal");
        let worksheetJarakRumah = workbook.addWorksheet("Jarak Rumah");
        let worksheetAlatTransportasi = workbook.addWorksheet("Alat Transportasi");
        let worksheetPenghasilan = workbook.addWorksheet("Penghasilan");

        //Data Siswa
        worksheet.columns = [
          { header: "NAMA", key: "name", width: 20 },
          { header: "EMAIL", key: "email", width: 20 },
          { header: "NIK SISWA", key: "nik_siswa", width: 20 },
          { header: "NISN", key: "nomor_induk", width: 20 },
          { header: "TANGGAL LAHIR", key: "tgl_lahir", width: 20 },
          { header: "TEMPAT", key: "tempat", width: 20 },
          { header: "JENIS KELAMIN", key: "jeniskelamin", width: 20 },
          { header: "AGAMA", key: "agama", width: 20 },
          { header: "ANAK KE", key: "anakke", width: 20 },
          { header: "JUMLAH SAUDARA", key: "jumlah_saudara", width: 20 },
          { header: "HOBI", key: "hobi", width: 20 },
          { header: "CITA-CITA", key: "cita_cita", width: 20 },
          { header: "JENJANG SEKOLAH", key: "jenjang", width: 20 },
          { header: "NAMA SEKOLAH", key: "nama_sekolah", width: 20 },
          { header: "STATUS SEKOLAH", key: "status_sekolah", width: 20 },
          { header: "NPSN", key: "npsn", width: 20 },
          { header: "ALAMAT SEKOLAH", key: "alamat_sekolah", width: 40 },
          { header: "KABUPATEN / KOTA SEKOLAH SEBELUMNYA", key: "kabkot_sekolah", width: 20 },
          { header: "NOMOR KK", key: "no_kk", width: 20 },
          { header: "NAMA KEPALA KELUARGA", key: "nama_kk", width: 20 },
          { header: "NIK AYAH", key: "nik_ayah", width: 20 },
          { header: "NAMA AYAH", key: "nama_ayah", width: 20 },
          { header: "TAHUN AYAH", key: "tahun_ayah", width: 20 },
          { header: "STATUS AYAH", key: "status_ayah", width: 20 },
          { header: "PENDIDIKAN AYAH", key: "pendidikan_ayah", width: 20 },
          { header: "PEKERJAAN AYAH", key: "pekerjaan_ayah", width: 20 },
          { header: "NO HANDPHONE AYAH", key: "telp_ayah", width: 20 },
          { header: "NIK IBU", key: "nik_ibu", width: 20 },
          { header: "NAMA IBU", key: "nama_ibu", width: 20 },
          { header: "TAHUN IBU", key: "tahun_ibu", width: 20 },
          { header: "STATUS IBU", key: "status_ibu", width: 20 },
          { header: "PENDIDIKAN IBU", key: "pendidikan_ibu", width: 20 },
          { header: "PEKERJAAN IBU", key: "pekerjaan_ibu", width: 20 },
          { header: "NO HANDPHONE IBU", key: "telp_ibu", width: 20 },
          { header: "TELEPON", key: "telp", width: 20 },
          { header: "ALAMAT", key: "alamat", width: 40 },
          { header: "PROVINSI", key: "provinsi", width: 20 },
          { header: "KABUPATEN / KOTA", key: "kabkota", width: 20 },
          { header: "KECAMATAN", key: "kecamatan", width: 20 },
          { header: "KELURAHAN", key: "kelurahan", width: 20 },
          { header: "KODE POS", key: "kode_pos", width: 20 },
          { header: "PENGHASILAN", key: "penghasilan", width: 20 },
        ];
        // const figureColumns = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18 ,19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42];
        // figureColumns.forEach((i) => {
        //   worksheet.getColumn(i).alignment = { horizontal: "left" };
        // });
        worksheet.autoFilter = 'A1:AP1';
        worksheet.addRows([{
          name: 'tes', 
          email: 'tes@gmail.com', 
          nik_siswa: '123', 
          nomor_induk: '123', 
          tgl_lahir: new Date(),
          tempat: 'Bogor', 
          jeniskelamin: 'Laki - Laki', 
          agama: 'Islam', 
          anakke: '1', 
          jumlah_saudara: '1', 
          hobi: '1', 
          cita_cita: '1', 
          jenjang: '1', 
          nama_sekolah: 'SD. Teka Teki', 
          status_sekolah: '1', 
          npsn: '123', 
          alamat_sekolah: 'Bogor', 
          kabkot_sekolah: '32.01', 
          no_kk: '123', 
          nama_kk: 'Andre', 
          nik_ayah: '123', 
          nama_ayah: 'Andre', 
          tahun_ayah: '1970', 
          status_ayah: '1', 
          pendidikan_ayah: '1', 
          pekerjaan_ayah: '1', 
          telp_ayah: '123456789', 
          nik_ibu: '123', 
          nama_ibu: 'Susi', 
          tahun_ibu: '1989', 
          status_ibu: '1', 
          pendidikan_ibu: '1', 
          pekerjaan_ibu: '1', 
          telp_ibu: '123456789', 
          telp: '123456789', 
          alamat: 'Bogor', 
          provinsi: '32', 
          kabkota: '32.01', 
          kecamatan: '32.01.01', 
          kelurahan: '32.01.01.1002', 
          kode_pos: '16913',
          penghasilan: '1',
        }]);
        worksheet.eachRow(function (row, rowNumber) {
          row.eachCell((cell, colNumber) => {
            if (rowNumber == 1) {
              cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: '2d9c5d' }
              }
              cell.border = {
                top: { style: 'medium' },
                left: { style: 'medium' },
                bottom: { style: 'medium' },
                right: { style: 'medium' }
              };
              cell.font = {
                bold: true,
                color: {
                  argb: 'ffffff'
                }
              };
              cell.alignment = {
                horizontal: 'center',
                vertical: 'middle',
                wrapText: true
              };
            }else{
              cell.font = {
                bold: false,
                color: {
                  argb: '000000'
                }
              };
              cell.alignment = {
                horizontal: 'left',
                vertical: 'middle',
                wrapText: true
              };
            }
          })
          row.commit();
        });
        
        //Pil Agama
        worksheetAgama.columns = [
          { header: "KODE", key: "kode", width: 15 },
          { header: "LABEL", key: "label", width: 15 }
        ];
        const figureColumnsAgama = [1, 2];
        figureColumnsAgama.forEach((i) => {
          worksheetAgama.getColumn(i).alignment = { horizontal: "left" };
        });
        worksheetAgama.addRows([
          { kode: 'Islam', label: 'Islam' },
          { kode: 'Katolik', label: 'Katolik' },
          { kode: 'Protestan', label: 'Protestan' },
          { kode: 'Hindu', label: 'Hindu' },
          { kode: 'Budha', label: 'Budha' }
        ]);

        //Pil Hobi
        worksheetHobi.columns = [
          { header: "KODE", key: "kode", width: 10 },
          { header: "LABEL", key: "label", width: 50 }
        ];
        const figureColumnsHobi = [1, 2];
        figureColumnsHobi.forEach((i) => {
          worksheetHobi.getColumn(i).alignment = { horizontal: "left" };
        });
        worksheetHobi.addRows([
          { kode: '1', label: 'Olahraga' },
          { kode: '2', label: 'Kesenian' },
          { kode: '3', label: 'Membaca' },
          { kode: '4', label: 'Menulis' },
          { kode: '5', label: 'Traveling' },
          { kode: '6', label: 'Lainnya' },
        ]);

        //Pil CitaCita
        worksheetCitaCita.columns = [
          { header: "KODE", key: "kode", width: 10 },
          { header: "LABEL", key: "label", width: 50 }
        ];
        const figureColumnsCitaCita = [1, 2];
        figureColumnsCitaCita.forEach((i) => {
          worksheetCitaCita.getColumn(i).alignment = { horizontal: "left" };
        });
        worksheetCitaCita.addRows([
          { kode: '1', label: 'PNS' },
          { kode: '2', label: 'TNI/PORLI' },
          { kode: '3', label: 'Guru/Dosen' },
          { kode: '4', label: 'Dokter' },
          { kode: '5', label: 'Politikus' },
          { kode: '6', label: 'Wiraswasta' },
          { kode: '7', label: 'Pekerja Seni/Lukis/Artis/Sejenis' },
          { kode: '8', label: 'Lainnya' },
        ]);

        //Pil JenjangSekolah
        worksheetJenjangSekolah.columns = [
          { header: "KODE", key: "kode", width: 10 },
          { header: "LABEL", key: "label", width: 50 }
        ];
        const figureColumnsJenjangSekolah = [1, 2];
        figureColumnsJenjangSekolah.forEach((i) => {
          worksheetJenjangSekolah.getColumn(i).alignment = { horizontal: "left" };
        });
        worksheetJenjangSekolah.addRows([
          { kode: '1', label: 'MI' },
          { kode: '2', label: 'SD' },
          { kode: '3', label: 'SD Terbuka' },
          { kode: '4', label: 'SLB-MI' },
          { kode: '5', label: 'Paket A' },
          { kode: '6', label: 'Salafiyah Ula' },
          { kode: '7', label: 'MU`adalah MI' },
          { kode: '8', label: 'SLB-SD' },
          { kode: '9', label: 'Lainnya' },
        ]);

        //Pil StatusSekolah
        worksheetStatusSekolah.columns = [
          { header: "KODE", key: "kode", width: 10 },
          { header: "LABEL", key: "label", width: 50 }
        ];
        const figureColumnsStatusSekolah = [1, 2];
        figureColumnsStatusSekolah.forEach((i) => {
          worksheetStatusSekolah.getColumn(i).alignment = { horizontal: "left" };
        });
        worksheetStatusSekolah.addRows([
          { kode: '1', label: 'Negeri' },
          { kode: '2', label: 'Swasta' },
        ]);

        //Pil StatusOrangTua
        worksheetStatusOrangTua.columns = [
          { header: "KODE", key: "kode", width: 10 },
          { header: "LABEL", key: "label", width: 50 }
        ];
        const figureColumnsStatusOrangTua = [1, 2];
        figureColumnsStatusOrangTua.forEach((i) => {
          worksheetStatusOrangTua.getColumn(i).alignment = { horizontal: "left" };
        });
        worksheetStatusOrangTua.addRows([
          { kode: '1', label: 'Masih Hidup' },
          { kode: '2', label: 'Sudah Mati' },
          { kode: '3', label: 'Tidak Diketahui' },
        ]);

        //Pil Pendidikan
        worksheetPendidikan.columns = [
          { header: "KODE", key: "kode", width: 10 },
          { header: "LABEL", key: "label", width: 50 }
        ];
        const figureColumnsPendidikan = [1, 2];
        figureColumnsPendidikan.forEach((i) => {
          worksheetPendidikan.getColumn(i).alignment = { horizontal: "left" };
        });
        worksheetPendidikan.addRows([
          { kode: '0', label: 'Tidak Berpendidikan Formal' },
          { kode: '1', label: 'SD/Sederajat' },
          { kode: '2', label: 'SMP/Sederajat' },
          { kode: '3', label: 'SMA/Sederajat' },
          { kode: '4', label: 'D1' },
          { kode: '5', label: 'D2' },
          { kode: '6', label: 'D3' },
          { kode: '7', label: 'S1' },
          { kode: '8', label: 'S2' },
          { kode: '9', label: '>S2' },
        ]);

        //Pil Pekerjaan
        worksheetPekerjaan.columns = [
          { header: "KODE", key: "kode", width: 10 },
          { header: "LABEL", key: "label", width: 50 }
        ];
        const figureColumnsPekerjaan = [1, 2];
        figureColumnsPekerjaan.forEach((i) => {
          worksheetPekerjaan.getColumn(i).alignment = { horizontal: "left" };
        });
        worksheetPekerjaan.addRows([
          { kode: '1', label: 'Tidak Bekerja' },
          { kode: '2', label: 'Pensiunan/Almarhum' },
          { kode: '3', label: 'PNS (selain Guru/Dosen/Dokter/Bidan/Perawat)' },
          { kode: '4', label: 'TNI/Polisi' },
          { kode: '5', label: 'Guru/Dosen' },
          { kode: '6', label: 'Pegawai Swasta' },
          { kode: '7', label: 'Pengusaha/Wiraswasta' },
          { kode: '8', label: 'Pengacara/Hakim/Jaksa/Notaris' },
          { kode: '9', label: 'Seniman/Pelukis/Artis/Sejenis' },
          { kode: '10', label: 'Dokter/Bidan/Perawat' },
          { kode: '11', label: 'Pilot/Pramugari' },
          { kode: '12', label: 'Pedagang' },
          { kode: '13', label: 'Petani/Peternak' },
          { kode: '14', label: 'Nelayan' },
          { kode: '15', label: 'Buruh (Tani/Pabrik/Bangunan)' },
          { kode: '16', label: 'Sopir/Masinis/Kondektur' },
          { kode: '17', label: 'Politikus' },
          { kode: '18', label: 'Lainnya' },
        ]);

        //Pil StatusTempatTinggal
        worksheetStatusTempatTinggal.columns = [
          { header: "KODE", key: "kode", width: 10 },
          { header: "LABEL", key: "label", width: 50 }
        ];
        const figureColumnsStatusTempatTinggal = [1, 2];
        figureColumnsStatusTempatTinggal.forEach((i) => {
          worksheetStatusTempatTinggal.getColumn(i).alignment = { horizontal: "left" };
        });
        worksheetStatusTempatTinggal.addRows([
          { kode: '1', label: 'Milik' },
          { kode: '2', label: 'Rumah Orangtua' },
          { kode: '3', label: 'Rumah Saudara/Kerabat' },
          { kode: '4', label: 'Rumah Dinas' },
        ]);

        //Pil JarakRumah
        worksheetJarakRumah.columns = [
          { header: "KODE", key: "kode", width: 10 },
          { header: "LABEL", key: "label", width: 50 }
        ];
        const figureColumnsJarakRumah = [1, 2];
        figureColumnsJarakRumah.forEach((i) => {
          worksheetJarakRumah.getColumn(i).alignment = { horizontal: "left" };
        });
        worksheetJarakRumah.addRows([
          { kode: '1', label: '< 1 Km' },
          { kode: '2', label: '1 - 3 Km' },
          { kode: '3', label: '3 - 5 Km' },
          { kode: '4', label: '5 - 10 Km' },
          { kode: '5', label: '> 10 Km' },
        ]);

        //Pil AlatTransportasi
        worksheetAlatTransportasi.columns = [
          { header: "KODE", key: "kode", width: 10 },
          { header: "LABEL", key: "label", width: 50 }
        ];
        const figureColumnsAlatTransportasi = [1, 2];
        figureColumnsAlatTransportasi.forEach((i) => {
          worksheetAlatTransportasi.getColumn(i).alignment = { horizontal: "left" };
        });
        worksheetAlatTransportasi.addRows([
          { kode: '1', label: 'Jalan Kaki' },
          { kode: '2', label: 'Sepeda' },
          { kode: '3', label: 'Sepeda Motor' },
          { kode: '4', label: 'Mobil Pribadi' },
          { kode: '5', label: 'Antar Jemput Sekolah' },
          { kode: '6', label: 'Angkutan Umum' },
          { kode: '7', label: 'Perahu/Sampan' },
          { kode: '8', label: 'Lainnya' },
        ]);

        //Pil Penghasilan
        worksheetPenghasilan.columns = [
          { header: "KODE", key: "kode", width: 10 },
          { header: "LABEL", key: "label", width: 50 }
        ];
        const figureColumnsPenghasilan = [1, 2];
        figureColumnsPenghasilan.forEach((i) => {
          worksheetPenghasilan.getColumn(i).alignment = { horizontal: "left" };
        });
        worksheetPenghasilan.addRows([
          { kode: '1', label: '<= Rp 500.000' },
          { kode: '2', label: 'Rp 500.001 - Rp 1.000.000' },
          { kode: '3', label: 'Rp 1.000.001 - Rp 2.000.000' },
          { kode: '4', label: 'Rp 2.000.001 - Rp 3.000.000' },
          { kode: '5', label: 'Rp 3.000.001 - Rp 5.000.000' },
          { kode: '6', label: '> Rp 5.000.000' },
        ]);

        res.setHeader(
          "Content-Disposition",
          "attachment; filename=TemplateDataSiswa.xlsx"
        );
      }else if(roleid === '2'){
        let worksheet = workbook.addWorksheet("Data Guru");
        let worksheetAgama = workbook.addWorksheet("Agama");
        let worksheetPendidikan = workbook.addWorksheet("Pendidikan");
        let worksheetJabatan = workbook.addWorksheet("Jabatan");
        let worksheetBidangMengajar = workbook.addWorksheet("Bidang Mengajar");

        //Data Guru
        worksheet.columns = [
          { header: "NAMA", key: "name", width: 20 },
          { header: "EMAIL", key: "email", width: 20 },
          { header: "TANGGAL LAHIR", key: "tgl_lahir", width: 20 },
          { header: "TEMPAT", key: "tempat", width: 20 },
          { header: "JENIS KELAMIN", key: "jeniskelamin", width: 20 },
          { header: "AGAMA", key: "agama", width: 20 },
          { header: "PENDIDIKAN TERAKHIR", key: "pendidikan_guru", width: 25 },
          { header: "JABATAN", key: "jabatan_guru", width: 20 },
          { header: "MENGAJAR BIDANG", key: "mengajar_bidang", width: 20 },
          { header: "MENGAJAR KELAS", key: "mengajar_kelas", width: 20 },
          { header: "TELEPON", key: "telp", width: 20 },
          { header: "ALAMAT", key: "alamat", width: 40 },
          { header: "PROVINSI", key: "provinsi", width: 20 },
          { header: "KABUPATEN / KOTA", key: "kabkota", width: 20 },
          { header: "KECAMATAN", key: "kecamatan", width: 20 },
          { header: "KELURAHAN", key: "kelurahan", width: 20 },
          { header: "KODE POS", key: "kode_pos", width: 20 },
        ];
        // const figureColumns = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17];
        // figureColumns.forEach((i) => {
        //   worksheet.getColumn(i).alignment = { horizontal: "left" };
        // });
        worksheet.autoFilter = 'A1:Q1';
        worksheet.addRows([{
          name: 'tes', 
          email: 'tes@gmail.com',
          tgl_lahir: new Date(),
          tempat: 'Bogor', 
          jeniskelamin: 'Laki - Laki', 
          agama: 'Islam',  
          pendidikan_guru: '5',  
          jabatan_guru: 'Staff TU',  
          mengajar_bidang: 'PKN',  
          mengajar_kelas: '7,8,9',  
          telp: '123456789', 
          alamat: 'Bogor', 
          provinsi: '32', 
          kabkota: '32.01', 
          kecamatan: '32.01.01', 
          kelurahan: '32.01.01.1002', 
          kode_pos: '16913',
        }]);
        worksheet.eachRow(function (row, rowNumber) {
          row.eachCell((cell, colNumber) => {
            if (rowNumber == 1) {
              cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: '2d9c5d' }
              }
              cell.border = {
                top: { style: 'medium' },
                left: { style: 'medium' },
                bottom: { style: 'medium' },
                right: { style: 'medium' }
              };
              cell.font = {
                bold: true,
                color: {
                  argb: 'ffffff'
                }
              };
              cell.alignment = {
                horizontal: 'center',
                vertical: 'middle',
                wrapText: true
              };
            }else{
              cell.font = {
                bold: false,
                color: {
                  argb: '000000'
                }
              };
              cell.alignment = {
                horizontal: 'left',
                vertical: 'middle',
                wrapText: true
              };
            }
          })
          row.commit();
        });

        //Pil Agama
        worksheetAgama.columns = [
          { header: "KODE", key: "kode", width: 15 },
          { header: "LABEL", key: "label", width: 15 }
        ];
        const figureColumnsAgama = [1, 2];
        figureColumnsAgama.forEach((i) => {
          worksheetAgama.getColumn(i).alignment = { horizontal: "left" };
        });
        worksheetAgama.addRows([
          { kode: 'Islam', label: 'Islam' },
          { kode: 'Katolik', label: 'Katolik' },
          { kode: 'Protestan', label: 'Protestan' },
          { kode: 'Hindu', label: 'Hindu' },
          { kode: 'Budha', label: 'Budha' }
        ]);

        //Pil Pendidikan
        worksheetPendidikan.columns = [
          { header: "KODE", key: "kode", width: 10 },
          { header: "LABEL", key: "label", width: 50 }
        ];
        const figureColumnsPendidikan = [1, 2];
        figureColumnsPendidikan.forEach((i) => {
          worksheetPendidikan.getColumn(i).alignment = { horizontal: "left" };
        });
        worksheetPendidikan.addRows([
          { kode: '0', label: 'Tidak Berpendidikan Formal' },
          { kode: '1', label: 'SD/Sederajat' },
          { kode: '2', label: 'SMP/Sederajat' },
          { kode: '3', label: 'SMA/Sederajat' },
          { kode: '4', label: 'D1' },
          { kode: '5', label: 'D2' },
          { kode: '6', label: 'D3' },
          { kode: '7', label: 'S1' },
          { kode: '8', label: 'S2' },
          { kode: '9', label: '>S2' },
        ]);

        //Pil Jabatan
        worksheetJabatan.columns = [
          { header: "KODE", key: "kode", width: 30 },
          { header: "LABEL", key: "label", width: 30 }
        ];
        const figureColumnsJabatan = [1, 2];
        figureColumnsJabatan.forEach((i) => {
          worksheetJabatan.getColumn(i).alignment = { horizontal: "left" };
        });
        worksheetJabatan.addRows([
          { value: 'Kepala Sekolah', label: 'Kepala Sekolah' },
          { value: 'WaKaBid. Kesiswaan', label: 'WaKaBid. Kesiswaan' },
          { value: 'WaKaBid. Kurikulum', label: 'WaKaBid. Kurikulum' },
          { value: 'WaKaBid. Sarpras', label: 'WaKaBid. Sarpras' },
          { value: 'Kepala TU', label: 'Kepala TU' },
          { value: 'Staff TU', label: 'Staff TU' },
          { value: 'Wali Kelas', label: 'Wali Kelas' },
          { value: 'BP / BK', label: 'BP / BK' },
          { value: 'Pembina Osis', label: 'Pembina Osis' },
          { value: 'Pembina Pramuka', label: 'Pembina Pramuka' },
          { value: 'Pembina Paskibra', label: 'Pembina Paskibra' },
        ]);

        //Pil Bidang Mengajar
        worksheetBidangMengajar.columns = [
          { header: "KODE", key: "kode", width: 30 },
          { header: "LABEL", key: "label", width: 30 }
        ];
        const figureColumnsBidangworksheetBidangMengajar = [1, 2];
        figureColumnsBidangworksheetBidangMengajar.forEach((i) => {
          worksheetBidangMengajar.getColumn(i).alignment = { horizontal: "left" };
        });
        worksheetBidangMengajar.addRows([
          { kode: 'Alquran Hadits', label: 'Alquran Hadits' },
          { kode: 'Aqidah Akhlak', label: 'Aqidah Akhlak' },
          { kode: 'Bahasa Arab', label: 'Bahasa Arab' },
          { kode: 'Bahasa Indonesia', label: 'Bahasa Indonesia' },
          { kode: 'Bahasa Inggris', label: 'Bahasa Inggris' },
          { kode: 'Bahasa Sunda', label: 'Bahasa Sunda' },
          { kode: 'BTQ', label: 'BTQ' },
          { kode: 'Fiqih', label: 'Fiqih' },
          { kode: 'IPA Terpadu', label: 'IPA Terpadu' },
          { kode: 'IPS Terpadu', label: 'IPS Terpadu' },
          { kode: 'Matematika', label: 'Matematika' },
          { kode: 'Penjasorkes', label: 'Penjasorkes' },
          { kode: 'PKN', label: 'PKN' },
          { kode: 'Prakarya', label: 'Prakarya' },
          { kode: 'Seni Budaya', label: 'Seni Budaya' },
          { kode: 'SKI', label: 'SKI' },
        ]);

        res.setHeader(
          "Content-Disposition",
          "attachment; filename=TemplateDataGuru.xlsx"
        );
      }
      
      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      );
  
      return workbook.xlsx.write(res).then(function () {
        res.status(200).end();
      });
    }catch(err){
      console.log(err)
      logger.error(`[ERROR ACCESS] --- Gagal memproses data !`);
      return response(res, { kode: 404, message: 'Gagal memproses data !' }, 404);
    }
  }  

  exportexcel = async (req, res, next) => {
    try{
      let query = { ...req.query }
      let params = { ...req.params }
      let cariData = {
        roleid: query.export === 'dariAdmin' ? params.cari : null, 
        kelas: query.export === 'dariAdmin' ? null : params.cari, 
        cetak: query.export === 'dariAdmin' ? params.cari : '3'
      }
      let attributes = []
      let include = {}
      let search = {
        or: {
          roleID: cariData.roleid,
          kelas: cariData.kelas
        },
        and: {},
      }
      let data = await this.readDataBy(attributes, search, include, 'viewusers.id', 'viewusers');
      if(!data.length) return response(res, { kode: 404, message: 'Data tidak di temukan !' }, 404);
  
      let tampung = []
      data.map((value) => {
        tampung = [
          ...tampung,
          {
            item_no: value.item_no,
            id: value.id,
            roleID: value.roleID,
            name: UpperFirstLetter(value.name),
            email: value.email,
            password: value.password,
            gambar: value.gambar,
            gambarGmail: value.gambarGmail,
            refresh_token: value.refresh_token,
            kodeOTP: value.kodeOTP,
            codeLog: value.codeLog,
            activeAkun: value.activeAkun,
            validasiAkun: value.validasiAkun,
            mutationAkun: value.mutationAkun,
            id_users_details: value.id_users_details,
            id_profile: value.id_profile,
            nik_siswa: value.nik_siswa,
            nomor_induk: value.nomor_induk,
            tgl_lahir: dateconvert(value.tgl_lahir),
            tempat: UpperFirstLetter(value.tempat),
            jeniskelamin: value.jeniskelamin,
            agama: value.agama,
            anakke: value.anakke,
            jumlah_saudara: value.jumlah_saudara,
            hobi: value.hobi,
            cita_cita: value.cita_cita,
            jenjang: value.jenjang,
            status_sekolah: value.status_sekolah,
            nama_sekolah: value.nama_sekolah ? UpperFirstLetter(value.nama_sekolah) : value.nama_sekolah,
            npsn: value.npsn,
            alamat_sekolah: value.alamat_sekolah ? UpperFirstLetter(value.alamat_sekolah) : value.alamat_sekolah,
            kabkot_sekolah: value.kabkot_sekolah,
            no_peserta_un: value.no_peserta_un,
            no_skhun: value.no_skhun,
            no_ijazah: value.no_ijazah,
            nilai_un: value.nilai_un,
            no_kk: value.no_kk,
            nama_kk: value.nama_kk ? UpperFirstLetter(value.nama_kk) : value.nama_kk,
            nama_ayah: value.nama_ayah ? UpperFirstLetter(value.nama_ayah) : value.nama_ayah,
            tahun_ayah: value.tahun_ayah,
            status_ayah: value.status_ayah,
            nik_ayah: value.nik_ayah,
            pendidikan_ayah: value.pendidikan_ayah,
            pekerjaan_ayah: value.pekerjaan_ayah,
            telp_ayah: value.telp_ayah,
            nama_ibu: value.nama_ibu ? UpperFirstLetter(value.nama_ibu) : value.nama_ibu,
            tahun_ibu: value.tahun_ibu,
            status_ibu: value.status_ibu,
            nik_ibu: value.nik_ibu,
            pendidikan_ibu: value.pendidikan_ibu,
            pekerjaan_ibu: value.pekerjaan_ibu,
            telp_ibu: value.telp_ibu,
            nama_wali: value.nama_wali ? UpperFirstLetter(value.nama_wali) : value.nama_wali,
            tahun_wali: value.tahun_wali,
            nik_wali: value.nik_wali,
            pendidikan_wali: value.pendidikan_wali,
            pekerjaan_wali: value.pekerjaan_wali,
            telp_wali: value.telp_wali,
            penghasilan: value.penghasilan,
            telp: value.telp,
            alamat: UpperFirstLetter(value.alamat),
            provinsi: value.provinsi,
            kabkota: value.kabkota,
            kecamatan: value.kecamatan,
            kelurahan: value.kelurahan,
            kode_pos: value.kode_pos,
            pendidikan_guru: value.pendidikan_guru,
            jabatan_guru: value.jabatan_guru,
            mengajar_bidang: value.mengajar_bidang,
            mengajar_kelas: value.mengajar_kelas,
            walikelas: value.walikelas,
            kelas: value.kelas,
            status_tempat_tinggal: value.status_tempat_tinggal,
            jarak_rumah: value.jarak_rumah,
            transportasi: value.transportasi,
            fc_ijazah: value.fc_ijazah,
            fc_skhun: value.fc_skhun,
            fc_kk: value.fc_kk,
            fc_ktp_ortu: value.fc_ktp_ortu,
            fc_akta_lahir: value.fc_akta_lahir,
            fc_skl: value.fc_skl,
            createdAt: value.createdAt,
            updatedAt: value.updatedAt,
            updatedUsers: value.updatedUsers,
            createdUsers: value.createdUsers,
            roleName: value.roleName,
            nama_provinsi: UpperFirstLetter(value.nama_provinsi),
            nama_kabkota: UpperFirstLetter(value.nama_kabkota),
            nama_kecamatan: UpperFirstLetter(value.kecamatan),
            nama_kelurahan: UpperFirstLetter(value.kelurahan),
            nama_kabkot_sekolah: value.nama_kabkot_sekolah ? UpperFirstLetter(value.nama_kabkot_sekolah) : value.nama_kabkot_sekolah 
          }
        ]
      })

      let workbook = new excel.Workbook();
      workbook.creator = 'Triyoga Ginanjar Pamungkas';
      workbook.created = new Date();
      if(cariData.cetak === '3'){
        let worksheet = workbook.addWorksheet("Data Siswa");
        let worksheetAgama = workbook.addWorksheet("Agama");
        let worksheetHobi = workbook.addWorksheet("Hobi");
        let worksheetCitaCita = workbook.addWorksheet("Cita - Cita");
        let worksheetJenjangSekolah = workbook.addWorksheet("Jenjang Sekolah");
        let worksheetStatusSekolah = workbook.addWorksheet("Status Sekolah");
        let worksheetStatusOrangTua = workbook.addWorksheet("Status Orang Tua");
        let worksheetPendidikan = workbook.addWorksheet("Pendidikan");
        let worksheetPekerjaan = workbook.addWorksheet("Pekerjaan");
        let worksheetStatusTempatTinggal = workbook.addWorksheet("Status Tempat Tinggal");
        let worksheetJarakRumah = workbook.addWorksheet("Jarak Rumah");
        let worksheetAlatTransportasi = workbook.addWorksheet("Alat Transportasi");
        let worksheetPenghasilan = workbook.addWorksheet("Penghasilan");

        //Data Siswa
        worksheet.columns = [
          { header: "NAMA", key: "name", width: 20 },
          { header: "EMAIL", key: "email", width: 20 },
          { header: "NIK SISWA", key: "nik_siswa", width: 20 },
          { header: "NISN", key: "nomor_induk", width: 20 },
          { header: "TANGGAL LAHIR", key: "tgl_lahir", width: 20 },
          { header: "TEMPAT", key: "tempat", width: 20 },
          { header: "JENIS KELAMIN", key: "jeniskelamin", width: 20 },
          { header: "AGAMA", key: "agama", width: 20 },
          { header: "ANAK KE", key: "anakke", width: 20 },
          { header: "JUMLAH SAUDARA", key: "jumlah_saudara", width: 20 },
          { header: "HOBI", key: "hobi", width: 20 },
          { header: "CITA-CITA", key: "cita_cita", width: 20 },
          { header: "JENJANG SEKOLAH", key: "jenjang", width: 20 },
          { header: "NAMA SEKOLAH", key: "nama_sekolah", width: 20 },
          { header: "STATUS SEKOLAH", key: "status_sekolah", width: 20 },
          { header: "NPSN", key: "npsn", width: 20 },
          { header: "ALAMAT SEKOLAH", key: "alamat_sekolah", width: 40 },
          { header: "KABUPATEN / KOTA SEKOLAH SEBELUMNYA", key: "kabkot_sekolah", width: 20 },
          { header: "NOMOR KK", key: "no_kk", width: 20 },
          { header: "NAMA KEPALA KELUARGA", key: "nama_kk", width: 20 },
          { header: "NIK AYAH", key: "nik_ayah", width: 20 },
          { header: "NAMA AYAH", key: "nama_ayah", width: 20 },
          { header: "TAHUN AYAH", key: "tahun_ayah", width: 20 },
          { header: "STATUS AYAH", key: "status_ayah", width: 20 },
          { header: "PENDIDIKAN AYAH", key: "pendidikan_ayah", width: 20 },
          { header: "PEKERJAAN AYAH", key: "pekerjaan_ayah", width: 20 },
          { header: "NO HANDPHONE AYAH", key: "telp_ayah", width: 20 },
          { header: "NIK IBU", key: "nik_ibu", width: 20 },
          { header: "NAMA IBU", key: "nama_ibu", width: 20 },
          { header: "TAHUN IBU", key: "tahun_ibu", width: 20 },
          { header: "STATUS IBU", key: "status_ibu", width: 20 },
          { header: "PENDIDIKAN IBU", key: "pendidikan_ibu", width: 20 },
          { header: "PEKERJAAN IBU", key: "pekerjaan_ibu", width: 20 },
          { header: "NO HANDPHONE IBU", key: "telp_ibu", width: 20 },
          { header: "NIK WALI", key: "nik_wali", width: 20 },
          { header: "NAMA WALI", key: "nama_wali", width: 20 },
          { header: "TAHUN WALI", key: "tahun_wali", width: 20 },
          { header: "PENDIDIKAN WALI", key: "pendidikan_wali", width: 20 },
          { header: "PEKERJAAN WALI", key: "pekerjaan_wali", width: 20 },
          { header: "NO HANDPHONE WALI", key: "telp_wali", width: 20 },
          { header: "TELEPON", key: "telp", width: 20 },
          { header: "ALAMAT", key: "alamat", width: 40 },
          { header: "PROVINSI", key: "provinsi", width: 20 },
          { header: "KABUPATEN / KOTA", key: "kabkota", width: 20 },
          { header: "KECAMATAN", key: "kecamatan", width: 20 },
          { header: "KELURAHAN", key: "kelurahan", width: 20 },
          { header: "KODE POS", key: "kode_pos", width: 20 },
          { header: "PENGHASILAN", key: "penghasilan", width: 20 },
          { header: "STATUS TEMPAT TINGGAL", key: "status_tempat_tinggal", width: 20 },
          { header: "JARAK RUMAH", key: "jarak_rumah", width: 20 },
          { header: "ALAT TRANSPORTASI", key: "transportasi", width: 20 },
          { header: "BERKAS IJAZAH", key: "fc_ijazah", width: 20 },
          { header: "BERKAS KARTU KELUARGA", key: "fc_kk", width: 20 },
          { header: "BERKAS KTP ORANG TUA", key: "fc_ktp_ortu", width: 20 },
          { header: "BERKAS AKTA LAHIR", key: "fc_akta_lahir", width: 20 },
          { header: "BERKAS SURAT KETERANGAN LULUS", key: "fc_skl", width: 20 },
        ];
        // const figureColumns = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 
        //   11, 12, 13, 14, 15, 16, 17, 18 ,19, 20, 
        //   21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 
        //   31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 
        //   41, 42, 43, 44, 45, 46, 47, 48, 49, 50,
        //   51, 52, 53, 54, 55, 56];
        // figureColumns.forEach((i) => {
        //   worksheet.getColumn(i).alignment = { horizontal: "left" };
        // });
        worksheet.autoFilter = 'A1:BD1';
        worksheet.addRows(tampung);
        worksheet.eachRow(function (row, rowNumber) {
          row.eachCell((cell, colNumber) => {
            if (rowNumber == 1) {
              cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: '2d9c5d' }
              }
              cell.border = {
                top: { style: 'medium' },
                left: { style: 'medium' },
                bottom: { style: 'medium' },
                right: { style: 'medium' }
              };
              cell.font = {
                bold: true,
                color: {
                  argb: 'ffffff'
                }
              };
              cell.alignment = {
                horizontal: 'center',
                vertical: 'middle',
                wrapText: true
              };
            }else{
              cell.border = {
                top: { style: 'dashDot' },
                left: { style: 'dashDot' },
                bottom: { style: 'dashDot' },
                right: { style: 'dashDot' }
              };
              cell.font = {
                bold: false,
                color: {
                  argb: '000000'
                }
              };
              cell.alignment = {
                horizontal: 'left',
                vertical: 'middle',
                wrapText: true
              };
            }
          })
          row.commit();
        });
        
        //Pil Agama
        worksheetAgama.columns = [
          { header: "KODE", key: "kode", width: 15 },
          { header: "LABEL", key: "label", width: 15 }
        ];
        const figureColumnsAgama = [1, 2];
        figureColumnsAgama.forEach((i) => {
          worksheetAgama.getColumn(i).alignment = { horizontal: "left" };
        });
        worksheetAgama.addRows([
          { kode: 'Islam', label: 'Islam' },
          { kode: 'Katolik', label: 'Katolik' },
          { kode: 'Protestan', label: 'Protestan' },
          { kode: 'Hindu', label: 'Hindu' },
          { kode: 'Budha', label: 'Budha' }
        ]);

        //Pil Hobi
        worksheetHobi.columns = [
          { header: "KODE", key: "kode", width: 10 },
          { header: "LABEL", key: "label", width: 50 }
        ];
        const figureColumnsHobi = [1, 2];
        figureColumnsHobi.forEach((i) => {
          worksheetHobi.getColumn(i).alignment = { horizontal: "left" };
        });
        worksheetHobi.addRows([
          { kode: '1', label: 'Olahraga' },
          { kode: '2', label: 'Kesenian' },
          { kode: '3', label: 'Membaca' },
          { kode: '4', label: 'Menulis' },
          { kode: '5', label: 'Traveling' },
          { kode: '6', label: 'Lainnya' },
        ]);

        //Pil CitaCita
        worksheetCitaCita.columns = [
          { header: "KODE", key: "kode", width: 10 },
          { header: "LABEL", key: "label", width: 50 }
        ];
        const figureColumnsCitaCita = [1, 2];
        figureColumnsCitaCita.forEach((i) => {
          worksheetCitaCita.getColumn(i).alignment = { horizontal: "left" };
        });
        worksheetCitaCita.addRows([
          { kode: '1', label: 'PNS' },
          { kode: '2', label: 'TNI/PORLI' },
          { kode: '3', label: 'Guru/Dosen' },
          { kode: '4', label: 'Dokter' },
          { kode: '5', label: 'Politikus' },
          { kode: '6', label: 'Wiraswasta' },
          { kode: '7', label: 'Pekerja Seni/Lukis/Artis/Sejenis' },
          { kode: '8', label: 'Lainnya' },
        ]);

        //Pil JenjangSekolah
        worksheetJenjangSekolah.columns = [
          { header: "KODE", key: "kode", width: 10 },
          { header: "LABEL", key: "label", width: 50 }
        ];
        const figureColumnsJenjangSekolah = [1, 2];
        figureColumnsJenjangSekolah.forEach((i) => {
          worksheetJenjangSekolah.getColumn(i).alignment = { horizontal: "left" };
        });
        worksheetJenjangSekolah.addRows([
          { kode: '1', label: 'MI' },
          { kode: '2', label: 'SD' },
          { kode: '3', label: 'SD Terbuka' },
          { kode: '4', label: 'SLB-MI' },
          { kode: '5', label: 'Paket A' },
          { kode: '6', label: 'Salafiyah Ula' },
          { kode: '7', label: 'MU`adalah MI' },
          { kode: '8', label: 'SLB-SD' },
          { kode: '9', label: 'Lainnya' },
        ]);

        //Pil StatusSekolah
        worksheetStatusSekolah.columns = [
          { header: "KODE", key: "kode", width: 10 },
          { header: "LABEL", key: "label", width: 50 }
        ];
        const figureColumnsStatusSekolah = [1, 2];
        figureColumnsStatusSekolah.forEach((i) => {
          worksheetStatusSekolah.getColumn(i).alignment = { horizontal: "left" };
        });
        worksheetStatusSekolah.addRows([
          { kode: '1', label: 'Negeri' },
          { kode: '2', label: 'Swasta' },
        ]);

        //Pil StatusOrangTua
        worksheetStatusOrangTua.columns = [
          { header: "KODE", key: "kode", width: 10 },
          { header: "LABEL", key: "label", width: 50 }
        ];
        const figureColumnsStatusOrangTua = [1, 2];
        figureColumnsStatusOrangTua.forEach((i) => {
          worksheetStatusOrangTua.getColumn(i).alignment = { horizontal: "left" };
        });
        worksheetStatusOrangTua.addRows([
          { kode: '1', label: 'Masih Hidup' },
          { kode: '2', label: 'Sudah Mati' },
          { kode: '3', label: 'Tidak Diketahui' },
        ]);

        //Pil Pendidikan
        worksheetPendidikan.columns = [
          { header: "KODE", key: "kode", width: 10 },
          { header: "LABEL", key: "label", width: 50 }
        ];
        const figureColumnsPendidikan = [1, 2];
        figureColumnsPendidikan.forEach((i) => {
          worksheetPendidikan.getColumn(i).alignment = { horizontal: "left" };
        });
        worksheetPendidikan.addRows([
          { kode: '0', label: 'Tidak Berpendidikan Formal' },
          { kode: '1', label: 'SD/Sederajat' },
          { kode: '2', label: 'SMP/Sederajat' },
          { kode: '3', label: 'SMA/Sederajat' },
          { kode: '4', label: 'D1' },
          { kode: '5', label: 'D2' },
          { kode: '6', label: 'D3' },
          { kode: '7', label: 'S1' },
          { kode: '8', label: 'S2' },
          { kode: '9', label: '>S2' },
        ]);

        //Pil Pekerjaan
        worksheetPekerjaan.columns = [
          { header: "KODE", key: "kode", width: 10 },
          { header: "LABEL", key: "label", width: 50 }
        ];
        const figureColumnsPekerjaan = [1, 2];
        figureColumnsPekerjaan.forEach((i) => {
          worksheetPekerjaan.getColumn(i).alignment = { horizontal: "left" };
        });
        worksheetPekerjaan.addRows([
          { kode: '1', label: 'Tidak Bekerja' },
          { kode: '2', label: 'Pensiunan/Almarhum' },
          { kode: '3', label: 'PNS (selain Guru/Dosen/Dokter/Bidan/Perawat)' },
          { kode: '4', label: 'TNI/Polisi' },
          { kode: '5', label: 'Guru/Dosen' },
          { kode: '6', label: 'Pegawai Swasta' },
          { kode: '7', label: 'Pengusaha/Wiraswasta' },
          { kode: '8', label: 'Pengacara/Hakim/Jaksa/Notaris' },
          { kode: '9', label: 'Seniman/Pelukis/Artis/Sejenis' },
          { kode: '10', label: 'Dokter/Bidan/Perawat' },
          { kode: '11', label: 'Pilot/Pramugari' },
          { kode: '12', label: 'Pedagang' },
          { kode: '13', label: 'Petani/Peternak' },
          { kode: '14', label: 'Nelayan' },
          { kode: '15', label: 'Buruh (Tani/Pabrik/Bangunan)' },
          { kode: '16', label: 'Sopir/Masinis/Kondektur' },
          { kode: '17', label: 'Politikus' },
          { kode: '18', label: 'Lainnya' },
        ]);

        //Pil StatusTempatTinggal
        worksheetStatusTempatTinggal.columns = [
          { header: "KODE", key: "kode", width: 10 },
          { header: "LABEL", key: "label", width: 50 }
        ];
        const figureColumnsStatusTempatTinggal = [1, 2];
        figureColumnsStatusTempatTinggal.forEach((i) => {
          worksheetStatusTempatTinggal.getColumn(i).alignment = { horizontal: "left" };
        });
        worksheetStatusTempatTinggal.addRows([
          { kode: '1', label: 'Milik' },
          { kode: '2', label: 'Rumah Orangtua' },
          { kode: '3', label: 'Rumah Saudara/Kerabat' },
          { kode: '4', label: 'Rumah Dinas' },
        ]);

        //Pil JarakRumah
        worksheetJarakRumah.columns = [
          { header: "KODE", key: "kode", width: 10 },
          { header: "LABEL", key: "label", width: 50 }
        ];
        const figureColumnsJarakRumah = [1, 2];
        figureColumnsJarakRumah.forEach((i) => {
          worksheetJarakRumah.getColumn(i).alignment = { horizontal: "left" };
        });
        worksheetJarakRumah.addRows([
          { kode: '1', label: '< 1 Km' },
          { kode: '2', label: '1 - 3 Km' },
          { kode: '3', label: '3 - 5 Km' },
          { kode: '4', label: '5 - 10 Km' },
          { kode: '5', label: '> 10 Km' },
        ]);

        //Pil AlatTransportasi
        worksheetAlatTransportasi.columns = [
          { header: "KODE", key: "kode", width: 10 },
          { header: "LABEL", key: "label", width: 50 }
        ];
        const figureColumnsAlatTransportasi = [1, 2];
        figureColumnsAlatTransportasi.forEach((i) => {
          worksheetAlatTransportasi.getColumn(i).alignment = { horizontal: "left" };
        });
        worksheetAlatTransportasi.addRows([
          { kode: '1', label: 'Jalan Kaki' },
          { kode: '2', label: 'Sepeda' },
          { kode: '3', label: 'Sepeda Motor' },
          { kode: '4', label: 'Mobil Pribadi' },
          { kode: '5', label: 'Antar Jemput Sekolah' },
          { kode: '6', label: 'Angkutan Umum' },
          { kode: '7', label: 'Perahu/Sampan' },
          { kode: '8', label: 'Lainnya' },
        ]);

        //Pil Penghasilan
        worksheetPenghasilan.columns = [
          { header: "KODE", key: "kode", width: 10 },
          { header: "LABEL", key: "label", width: 50 }
        ];
        const figureColumnsPenghasilan = [1, 2];
        figureColumnsPenghasilan.forEach((i) => {
          worksheetPenghasilan.getColumn(i).alignment = { horizontal: "left" };
        });
        worksheetPenghasilan.addRows([
          { kode: '1', label: '<= Rp 500.000' },
          { kode: '2', label: 'Rp 500.001 - Rp 1.000.000' },
          { kode: '3', label: 'Rp 1.000.001 - Rp 2.000.000' },
          { kode: '4', label: 'Rp 2.000.001 - Rp 3.000.000' },
          { kode: '5', label: 'Rp 3.000.001 - Rp 5.000.000' },
          { kode: '6', label: '> Rp 5.000.000' },
        ]);

        res.setHeader(
          "Content-Disposition",
          "attachment; filename=Data Siswa.xlsx"
        );
      }else if(cariData.cetak === '2'){
        let worksheet = workbook.addWorksheet("Data Guru");
        let worksheetAgama = workbook.addWorksheet("Agama");
        let worksheetPendidikan = workbook.addWorksheet("Pendidikan");
        let worksheetJabatan = workbook.addWorksheet("Jabatan");
        let worksheetBidangMengajar = workbook.addWorksheet("Bidang Mengajar");

        //Data Guru
        worksheet.columns = [
          { header: "NAMA", key: "name", width: 20 },
          { header: "EMAIL", key: "email", width: 20 },
          { header: "TANGGAL LAHIR", key: "tgl_lahir", width: 20 },
          { header: "TEMPAT", key: "tempat", width: 20 },
          { header: "JENIS KELAMIN", key: "jeniskelamin", width: 20 },
          { header: "AGAMA", key: "agama", width: 20 },
          { header: "PENDIDIKAN TERAKHIR", key: "pendidikan_guru", width: 25 },
          { header: "JABATAN", key: "jabatan_guru", width: 20 },
          { header: "MENGAJAR BIDANG", key: "mengajar_bidang", width: 20 },
          { header: "MENGAJAR KELAS", key: "mengajar_kelas", width: 20 },
          { header: "WALI KELAS", key: "walikelas", width: 20 },
          { header: "TELEPON", key: "telp", width: 20 },
          { header: "ALAMAT", key: "alamat", width: 40 },
          { header: "PROVINSI", key: "provinsi", width: 20 },
          { header: "KABUPATEN / KOTA", key: "kabkota", width: 20 },
          { header: "KECAMATAN", key: "kecamatan", width: 20 },
          { header: "KELURAHAN", key: "kelurahan", width: 20 },
          { header: "KODE POS", key: "kode_pos", width: 20 },
        ];
        // const figureColumns = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18];
        // figureColumns.forEach((i) => {
        //   worksheet.getColumn(i).alignment = { horizontal: "left" };
        // });
        worksheet.autoFilter = 'A1:R1';
        worksheet.addRows(tampung);
        worksheet.eachRow(function (row, rowNumber) {
          row.eachCell((cell, colNumber) => {
            if (rowNumber == 1) {
              cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: '2d9c5d' }
              }
              cell.border = {
                top: { style: 'medium' },
                left: { style: 'medium' },
                bottom: { style: 'medium' },
                right: { style: 'medium' }
              };
              cell.font = {
                bold: true,
                color: {
                  argb: 'ffffff'
                }
              };
              cell.alignment = {
                horizontal: 'center',
                vertical: 'middle',
                wrapText: true
              };
            }else{
              cell.border = {
                top: { style: 'dashDot' },
                left: { style: 'dashDot' },
                bottom: { style: 'dashDot' },
                right: { style: 'dashDot' }
              };
              cell.font = {
                bold: false,
                color: {
                  argb: '000000'
                }
              };
              cell.alignment = {
                horizontal: 'left',
                vertical: 'middle',
                wrapText: true
              };
            }
          })
          row.commit();
        });

        //Pil Agama
        worksheetAgama.columns = [
          { header: "KODE", key: "kode", width: 15 },
          { header: "LABEL", key: "label", width: 15 }
        ];
        const figureColumnsAgama = [1, 2];
        figureColumnsAgama.forEach((i) => {
          worksheetAgama.getColumn(i).alignment = { horizontal: "left" };
        });
        worksheetAgama.addRows([
          { kode: 'Islam', label: 'Islam' },
          { kode: 'Katolik', label: 'Katolik' },
          { kode: 'Protestan', label: 'Protestan' },
          { kode: 'Hindu', label: 'Hindu' },
          { kode: 'Budha', label: 'Budha' }
        ]);

        //Pil Pendidikan
        worksheetPendidikan.columns = [
          { header: "KODE", key: "kode", width: 10 },
          { header: "LABEL", key: "label", width: 50 }
        ];
        const figureColumnsPendidikan = [1, 2];
        figureColumnsPendidikan.forEach((i) => {
          worksheetPendidikan.getColumn(i).alignment = { horizontal: "left" };
        });
        worksheetPendidikan.addRows([
          { kode: '0', label: 'Tidak Berpendidikan Formal' },
          { kode: '1', label: 'SD/Sederajat' },
          { kode: '2', label: 'SMP/Sederajat' },
          { kode: '3', label: 'SMA/Sederajat' },
          { kode: '4', label: 'D1' },
          { kode: '5', label: 'D2' },
          { kode: '6', label: 'D3' },
          { kode: '7', label: 'S1' },
          { kode: '8', label: 'S2' },
          { kode: '9', label: '>S2' },
        ]);

        //Pil Jabatan
        worksheetJabatan.columns = [
          { header: "KODE", key: "kode", width: 30 },
          { header: "LABEL", key: "label", width: 30 }
        ];
        const figureColumnsJabatan = [1, 2];
        figureColumnsJabatan.forEach((i) => {
          worksheetJabatan.getColumn(i).alignment = { horizontal: "left" };
        });
        worksheetJabatan.addRows([
          { value: 'Kepala Sekolah', label: 'Kepala Sekolah' },
          { value: 'WaKaBid. Kesiswaan', label: 'WaKaBid. Kesiswaan' },
          { value: 'WaKaBid. Kurikulum', label: 'WaKaBid. Kurikulum' },
          { value: 'WaKaBid. Sarpras', label: 'WaKaBid. Sarpras' },
          { value: 'Kepala TU', label: 'Kepala TU' },
          { value: 'Staff TU', label: 'Staff TU' },
          { value: 'Wali Kelas', label: 'Wali Kelas' },
          { value: 'BP / BK', label: 'BP / BK' },
          { value: 'Pembina Osis', label: 'Pembina Osis' },
          { value: 'Pembina Pramuka', label: 'Pembina Pramuka' },
          { value: 'Pembina Paskibra', label: 'Pembina Paskibra' },
        ]);

        //Pil Bidang Mengajar
        worksheetBidangMengajar.columns = [
          { header: "KODE", key: "kode", width: 30 },
          { header: "LABEL", key: "label", width: 30 }
        ];
        const figureColumnsBidangworksheetBidangMengajar = [1, 2];
        figureColumnsBidangworksheetBidangMengajar.forEach((i) => {
          worksheetBidangMengajar.getColumn(i).alignment = { horizontal: "left" };
        });
        worksheetBidangMengajar.addRows([
          { kode: 'Alquran Hadits', label: 'Alquran Hadits' },
          { kode: 'Aqidah Akhlak', label: 'Aqidah Akhlak' },
          { kode: 'Bahasa Arab', label: 'Bahasa Arab' },
          { kode: 'Bahasa Indonesia', label: 'Bahasa Indonesia' },
          { kode: 'Bahasa Inggris', label: 'Bahasa Inggris' },
          { kode: 'Bahasa Sunda', label: 'Bahasa Sunda' },
          { kode: 'BTQ', label: 'BTQ' },
          { kode: 'Fiqih', label: 'Fiqih' },
          { kode: 'IPA Terpadu', label: 'IPA Terpadu' },
          { kode: 'IPS Terpadu', label: 'IPS Terpadu' },
          { kode: 'Matematika', label: 'Matematika' },
          { kode: 'Penjasorkes', label: 'Penjasorkes' },
          { kode: 'PKN', label: 'PKN' },
          { kode: 'Prakarya', label: 'Prakarya' },
          { kode: 'Seni Budaya', label: 'Seni Budaya' },
          { kode: 'SKI', label: 'SKI' },
        ]);

        res.setHeader(
          "Content-Disposition",
          "attachment; filename=Data Guru.xlsx"
        );
      }

      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      );
    
      return workbook.xlsx.write(res).then(function () {
        res.status(200).end();
      });
    }catch(err){
      console.log(err)
      logger.error(`[ERROR ACCESS] --- Gagal memproses data !`);
      return response(res, { kode: 404, message: 'Gagal memproses data !' }, 404);
    }
  }

  importexcel = async (req, res, next) => {
    try{
      const { body, files } = req;
      const dir=files[0];
      let jsonData = [];
      readXlsxFile(dir.path, { getSheets: true }).then((sheets) => {
        let sheet
        sheets.forEach((obj) => {
          if(obj.name == 'Data Siswa') return sheet = obj.name
        });
        if(sheet == 'Data Siswa'){
          readXlsxFile(dir.path).then(async (rows) => {
            rows.shift();
            rows.forEach((row) => {
              let data = {
                name: row[0], 
                email: row[1], 
                nik_siswa: row[2], 
                nomor_induk: row[3], 
                tgl_lahir: convertDate(row[4]),
                tempat: row[5], 
                jeniskelamin: row[6], 
                agama: row[7], 
                anakke: row[8], 
                jumlah_saudara: row[9], 
                hobi: row[10], 
                cita_cita: row[11], 
                jenjang: row[12], 
                nama_sekolah: row[13], 
                status_sekolah: row[14], 
                npsn: row[15], 
                alamat_sekolah: row[16], 
                kabkot_sekolah: row[17], 
                no_kk: row[18], 
                nama_kk: row[19], 
                nik_ayah: row[20], 
                nama_ayah: row[21], 
                tahun_ayah: row[22], 
                status_ayah: row[23], 
                pendidikan_ayah: row[24], 
                pekerjaan_ayah: row[25], 
                telp_ayah: row[26], 
                nik_ibu: row[27], 
                nama_ibu: row[28], 
                tahun_ibu: row[29], 
                status_ibu: row[30], 
                pendidikan_ibu: row[31], 
                pekerjaan_ibu: row[32], 
                telp_ibu: row[33], 
                telp: row[34], 
                alamat: row[35], 
                provinsi: row[36], 
                kabkota: row[37], 
                kecamatan: row[38], 
                kelurahan: row[39], 
                kode_pos: row[40],
                penghasilan: row[41],
              };
              jsonData.push(data);
            });
            await this.prosesImport(res, jsonData)
          });
        }else{
          return response(res, { kode: 404, message: 'File yang di imoport tidak sesuai dengan format, cek kembali file yang di import !' }, 404);
        }
      });
    }catch(err){
      console.log(err)
      logger.error(`[ERROR ACCESS] --- Gagal memproses data !`);
      return response(res, { kode: 404, message: 'Gagal memproses data !' }, 404);
    }
  }

  prosesImport = async (res, jsonData = []) => {
    let kumpulInsert = [], kumpulNotInsert = []
    for (const [k, v] of Object.entries(jsonData)) {
      let attributes = []
      let include = {}
      let search = {
        or: {
          email: v.email,
          nomor_induk: v.nomor_induk
        },
        and: {},
      }
      let readData = await this.readDataBy(attributes, search, include, 'viewusers.id', 'viewusers');
      if(readData.length) {
        kumpulNotInsert.push(readData[0])
      }else{
        jsonData.map(row => {
          if(row.email == v.email) return kumpulInsert.push(row)
        });
      }
    }
    await this.prosesNotImport(res, kumpulNotInsert)
    // if(!kumpulInsert.length) return response(res, { kode: 404, message: 'data yang ingin di import sudah tersedia di database !' }, 404);
    for (const [k, v] of Object.entries(kumpulInsert)) {
      let kodeOTP = makeRandom(8)
      let salt = await bcrypt.genSalt();
      let hashPassword = await bcrypt.hash(kodeOTP, salt);
      let kirimdata1 = {
        name: v.name,
        email: v.email,
        roleID: '3',
        password: hashPassword,
        activeAkun: '1',
        kodeOTP: kodeOTP
      }
      let insertUsers = await MainModel.insert(kirimdata1, 'users');
      if(!insertUsers) { await MainModel.delete({ email: v.email }, 'users'); }
      let search2 = {
        or: {},
        and: {
          email: v.email
        },
      }
      let readData2 = await this.readDataBy(attributes, search2, include, 'users.id', 'users');
      let kirimdata2 = {
        id_profile: readData2[0].id,
        nomor_induk: v.nomor_induk,
        nik_siswa: v.nik_siswa,
        tempat: v.tempat,
        tgl_lahir: v.tgl_lahir,
        jeniskelamin: v.jeniskelamin,
        agama: v.agama,
        telp: v.telp,
        alamat: v.alamat,
        provinsi: v.provinsi,
        kabkota: v.kabkota,
        kecamatan: v.kecamatan,
        kelurahan: v.kelurahan,
        kode_pos: v.kode_pos,
        anakke: v.anakke,
        jumlah_saudara: v.jumlah_saudara,
        hobi: v.hobi,
        cita_cita: v.cita_cita,
        jenjang: v.jenjang,
        status_sekolah: v.status_sekolah,
        nama_sekolah: v.nama_sekolah,
        npsn: v.npsn,
        alamat_sekolah: v.alamat_sekolah,
        kabkot_sekolah: v.kabkot_sekolah,
        no_kk: v.no_kk,
        nama_kk: v.nama_kk,
        penghasilan: v.penghasilan,
        nik_ayah: v.nik_ayah,
        nama_ayah: v.nama_ayah,
        tahun_ayah: v.tahun_ayah,
        status_ayah: v.status_ayah,
        pendidikan_ayah: v.pendidikan_ayah,
        pekerjaan_ayah: v.pekerjaan_ayah,
        telp_ayah: v.telp_ayah,
        nik_ibu: v.nik_ibu,
        nama_ibu: v.nama_ibu,
        tahun_ibu: v.tahun_ibu,
        status_ibu: v.status_ibu,
        pendidikan_ibu: v.pendidikan_ibu,
        pekerjaan_ibu: v.pekerjaan_ibu,
        telp_ibu: v.telp_ibu,
      }
      await MainModel.insert(kirimdata2, 'users_details');
    }
  }

  prosesNotImport = async (res, jsonData = []) => {
    // console.log(jsonData)
    // process.exit()
    let workbook = new excel.Workbook();
    workbook.creator = 'Triyoga Ginanjar Pamungkas';
    workbook.created = new Date();

    let worksheet = workbook.addWorksheet("Data Siswa");

    //Data Siswa
    worksheet.columns = [
      { header: "NAMA", key: "name", width: 20 },
      { header: "EMAIL", key: "email", width: 20 },
      { header: "NIK SISWA", key: "nik_siswa", width: 20 },
      { header: "NISN", key: "nomor_induk", width: 20 },
      { header: "TANGGAL LAHIR", key: "tgl_lahir", width: 20 },
      { header: "TEMPAT", key: "tempat", width: 20 },
      { header: "JENIS KELAMIN", key: "jeniskelamin", width: 20 },
      { header: "AGAMA", key: "agama", width: 20 },
      { header: "ANAK KE", key: "anakke", width: 20 },
      { header: "JUMLAH SAUDARA", key: "jumlah_saudara", width: 20 },
      { header: "HOBI", key: "hobi", width: 20 },
      { header: "CITA-CITA", key: "cita_cita", width: 20 },
      { header: "JENJANG SEKOLAH", key: "jenjang", width: 20 },
      { header: "NAMA SEKOLAH", key: "nama_sekolah", width: 20 },
      { header: "STATUS SEKOLAH", key: "status_sekolah", width: 20 },
      { header: "NPSN", key: "npsn", width: 20 },
      { header: "ALAMAT SEKOLAH", key: "alamat_sekolah", width: 40 },
      { header: "KABUPATEN / KOTA SEKOLAH SEBELUMNYA", key: "kabkot_sekolah", width: 20 },
      { header: "NOMOR KK", key: "no_kk", width: 20 },
      { header: "NAMA KEPALA KELUARGA", key: "nama_kk", width: 20 },
      { header: "NIK AYAH", key: "nik_ayah", width: 20 },
      { header: "NAMA AYAH", key: "nama_ayah", width: 20 },
      { header: "TAHUN AYAH", key: "tahun_ayah", width: 20 },
      { header: "STATUS AYAH", key: "status_ayah", width: 20 },
      { header: "PENDIDIKAN AYAH", key: "pendidikan_ayah", width: 20 },
      { header: "PEKERJAAN AYAH", key: "pekerjaan_ayah", width: 20 },
      { header: "NO HANDPHONE AYAH", key: "telp_ayah", width: 20 },
      { header: "NIK IBU", key: "nik_ibu", width: 20 },
      { header: "NAMA IBU", key: "nama_ibu", width: 20 },
      { header: "TAHUN IBU", key: "tahun_ibu", width: 20 },
      { header: "STATUS IBU", key: "status_ibu", width: 20 },
      { header: "PENDIDIKAN IBU", key: "pendidikan_ibu", width: 20 },
      { header: "PEKERJAAN IBU", key: "pekerjaan_ibu", width: 20 },
      { header: "NO HANDPHONE IBU", key: "telp_ibu", width: 20 },
      { header: "NIK WALI", key: "nik_wali", width: 20 },
      { header: "NAMA WALI", key: "nama_wali", width: 20 },
      { header: "TAHUN WALI", key: "tahun_wali", width: 20 },
      { header: "PENDIDIKAN WALI", key: "pendidikan_wali", width: 20 },
      { header: "PEKERJAAN WALI", key: "pekerjaan_wali", width: 20 },
      { header: "NO HANDPHONE WALI", key: "telp_wali", width: 20 },
      { header: "TELEPON", key: "telp", width: 20 },
      { header: "ALAMAT", key: "alamat", width: 40 },
      { header: "PROVINSI", key: "provinsi", width: 20 },
      { header: "KABUPATEN / KOTA", key: "kabkota", width: 20 },
      { header: "KECAMATAN", key: "kecamatan", width: 20 },
      { header: "KELURAHAN", key: "kelurahan", width: 20 },
      { header: "KODE POS", key: "kode_pos", width: 20 },
      { header: "PENGHASILAN", key: "penghasilan", width: 20 },
      { header: "STATUS TEMPAT TINGGAL", key: "status_tempat_tinggal", width: 20 },
      { header: "JARAK RUMAH", key: "jarak_rumah", width: 20 },
      { header: "ALAT TRANSPORTASI", key: "transportasi", width: 20 },
      { header: "BERKAS IJAZAH", key: "fc_ijazah", width: 20 },
      { header: "BERKAS KARTU KELUARGA", key: "fc_kk", width: 20 },
      { header: "BERKAS KTP ORANG TUA", key: "fc_ktp_ortu", width: 20 },
      { header: "BERKAS AKTA LAHIR", key: "fc_akta_lahir", width: 20 },
      { header: "BERKAS SURAT KETERANGAN LULUS", key: "fc_skl", width: 20 },
    ];

    worksheet.autoFilter = 'A1:BD1';
    worksheet.addRows(jsonData);
    worksheet.eachRow(function (row, rowNumber) {
      row.eachCell((cell, colNumber) => {
        if (rowNumber == 1) {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '2d9c5d' }
          }
          cell.border = {
            top: { style: 'medium' },
            left: { style: 'medium' },
            bottom: { style: 'medium' },
            right: { style: 'medium' }
          };
          cell.font = {
            bold: true,
            color: {
              argb: 'ffffff'
            }
          };
          cell.alignment = {
            horizontal: 'center',
            vertical: 'middle',
            wrapText: true
          };
        }else{
          cell.border = {
            top: { style: 'dashDot' },
            left: { style: 'dashDot' },
            bottom: { style: 'dashDot' },
            right: { style: 'dashDot' }
          };
          cell.font = {
            bold: false,
            color: {
              argb: '000000'
            }
          };
          cell.alignment = {
            horizontal: 'left',
            vertical: 'middle',
            wrapText: true
          };
        }
      })
      row.commit();
    });

    let path = "./src/public/excel";
    let hasil = await workbook.xlsx.writeFile(`${path}/Data Siswa (not import).xlsx`)
    .then(() => {
      res.send({
        kode: 200,
        message: 'Berhasil',
        path: `${path}/Data Siswa (not import).xlsx`,
      });
    });
  }
}

module.exports = new MainController;