//TODO: FOR SOME REASON rIDs defined in presentation_rels are used in the slide.xml files so we need to clean the deleted an dupdate the indexdes
import JSZip, { JSZipObject } from "jszip";
const { parseStringPromise, Builder } = require('xml2js');
import fs from "fs";
import path from "path"
const outpath = './test-out'

function deleteObjectWithProperty(obj, targetKey, targetValue) {
    if (typeof obj !== 'object' || obj === null) {
        return obj;
    }

    if (Array.isArray(obj)) {
        for (let i = 0; i < obj.length; i++) {
            if (typeof obj[i] === 'object' && obj[i] !== null && obj[i][targetKey] === targetValue) {
                obj.splice(i, 1);
                i--;
            } else {
                obj[i] = deleteObjectWithProperty(obj[i], targetKey, targetValue);
            }
        }
    } else {
        for (const key in obj) {
            if (obj.hasOwnProperty(key)) {
                if (typeof obj[key] === 'object' && obj[key] !== null && obj[key][targetKey] === targetValue) {
                    delete obj[key];
                } else {
                    obj[key] = deleteObjectWithProperty(obj[key], targetKey, targetValue);
                }
            }
        }
    }

    return obj;
}

export class PowerPointEditor {
    public powerpointData: JSZip = new JSZip();
    public async loadPowerPoint() {
        const powerpoint: Buffer = fs.readFileSync("./test.pptx");
        const zip = new JSZip();
        this.powerpointData = await zip.loadAsync(powerpoint);       // Extract each file from the zip
    }
    async savePowerPoint() {

        return new Promise((resolve, reject) => {
            try {
                if (!this.powerpointData) {
                    throw new Error("No powerpoint data")
                }
                this.powerpointData.generateNodeStream({ streamFiles: true })
                    .pipe(fs.createWriteStream(`${outpath}/test-powerpoint.pptx`)
                        .on("finish", () => {
                            resolve('Done');
                        }));
            }
            catch (error) {
                reject(error);
            }
        });

    }
    public async unzipPowerPoint() {
        const powerpoint: Buffer = fs.readFileSync("./test.pptx");
        const zip = new JSZip();
        this.powerpointData = await zip.loadAsync(powerpoint);       // Extract each file from the zip
        for (const filename in zip.files) {
            const file = zip.files[filename];
            if (!file.dir) {
                // Extract the file contents
                const content: Buffer = await file.async("nodebuffer");
                const filePath = path.join(outpath, "powerpoint", filename);
                const directory = path.dirname(filePath);
                fs.mkdirSync(directory, { recursive: true });
                fs.writeFileSync(`${outpath}/powerpoint/${filename}`, content);
            }
        }
        console.log("DONE UNZIPPED");
    }


    public zipPowerPoint(): Promise<Buffer> {
        const allFilePaths: string[] = fs.readdirSync(`${outpath}/powerpoint/`, { recursive: true }) as string[];
        console.log(allFilePaths);
        const zip = new JSZip();
        for (const file of allFilePaths) {
            const stat = fs.lstatSync(`${outpath}/powerpoint/${file}`);
            if (stat.isDirectory()) {
                continue;
            }
            zip.file(file, fs.readFileSync(`${outpath}/powerpoint/${file}`));
        }
        return new Promise((resolve, reject) => {
            try {
                zip.generateNodeStream({ streamFiles: true })
                    .pipe(fs.createWriteStream(`${outpath}/powerpoint.pptx`)
                        .on("finish", () => {
                            resolve(fs.readFileSync(`${outpath}/powerpoint.pptx`));
                        }));
            }
            catch (error) {
                reject(error);
            }
        });
    }
    public writePowerPoint(powerpoint: Buffer, filePath: string) {
        fs.writeFileSync(filePath, powerpoint)

    }
    public async cleanup() {
        fs.rmSync(`${outpath}/powerpoint`, { recursive: true, force: true });
        //Possibly just delete whole ./out/
    }


    public async deleteSlide(index: number) {
        const slideXmlPath = `ppt/slides/slide${index}.xml`;
        this.powerpointData.remove(slideXmlPath)
        this.powerpointData.remove(`ppt/slides/_rels/slide${index}.xml.rels`)
        const presPath = 'ppt/presentation.xml';
        const relsPath = 'ppt/_rels/presentation.xml.rels';
        const relsData = await this.powerpointData.file(relsPath)?.async('string');
        if (relsData) {
            var relsXml = await parseStringPromise(relsData);
            //Delete all refs to slide we removed
            var rId = 0
            relsXml.Relationships.Relationship.forEach(rel => {
                if (!(rel['$'].Target.includes(`slide${index}.xml`))) {
                    return true
                }
                else {
                    rId = rel['$'].Id
                    return false
                }
            }
            );
            var presentationData = await this.powerpointData.file(presPath)?.async('string');
            var presXML = await parseStringPromise(presentationData)
            presXML['p:presentation']['p:sldIdLst'][0]['p:sldId'] = presXML['p:presentation']['p:sldIdLst'][0]['p:sldId'].filter((rel) => {
                if (!(rel['$']['r:id'] == rId)) {
                    return true;
                }
                else {
                    console.log("ID was found removing")
                    return false
                }
            })
            var builder = new Builder();
            const updatedPresXML = builder.buildObject(presXML);
            this.powerpointData.file(presPath, updatedPresXML);
            // Convert back to XML string and update in zip
            var builder = new Builder();
            const updatedRelsXml = builder.buildObject(relsXml);
            //const updatedRelsXml = xml2js.Builder(relsXml, { compact: true });
            this.powerpointData.file(relsPath, updatedRelsXml);

        } else {
            throw new Error('Failed to read presentation.xml.rels');
        }
    }
}


const test = async function() {
    const pe = new PowerPointEditor();
    await pe.loadPowerPoint()
    await pe.deleteSlide(1)
    await pe.savePowerPoint()
}

test()
