import { Upload, Checkbox, Banner, Tabs, TabPane } from "@douyinfe/semi-ui";
import { STATUS } from "../../../data/constants";
import { useTranslation } from "react-i18next";
import CodeEditor from "../../CodeEditor";

function fileExt(name) {
  if (!name) return "";
  const m = String(name).toLowerCase().match(/\.([^.]+)$/);
  return m ? m[1] : "";
}

export default function ImportSource({
  importData,
  setImportData,
  error,
  setError,
}) {
  const { t } = useTranslation();

  return (
    <div>
      <Tabs>
        <TabPane tab={t("insert_sql")} itemKey="text-import">
          <CodeEditor
            height={224}
            language="sql"
            onChange={(value) => {
              setImportData((prev) => ({ ...prev, src: value, srcType: "sql" }));
              setError({
                type: STATUS.NONE,
                message: "",
              });
            }}
          />
        </TabPane>
        <TabPane tab={t("upload_file")} itemKey="file-import">
          <Upload
            action="#"
            beforeUpload={({ file, fileList }) => {
              const f = fileList[0].fileInstance;
              if (!f) {
                return;
              }
              const ext = fileExt(f.name);

              if (ext === "xlsx" || ext === "xls") {
                // xlsx → ArrayBuffer → base64
                const reader = new FileReader();
                reader.onload = async (e) => {
                  const arrayBuffer = e.target.result;
                  // 转 base64 便于存 React state（不放对象）
                  const bytes = new Uint8Array(arrayBuffer);
                  let bin = "";
                  for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
                  const b64 = btoa(bin);
                  setImportData((prev) => ({
                    ...prev,
                    src: "",
                    srcType: "xlsx",
                    xlsxB64: b64,
                  }));
                };
                reader.readAsArrayBuffer(f);
              } else {
                // 默认按 sql 文本处理
                const reader = new FileReader();
                reader.onload = async (e) => {
                  setImportData((prev) => ({
                    ...prev,
                    src: e.target.result,
                    srcType: "sql",
                    xlsxB64: null,
                  }));
                };
                reader.readAsText(f);
              }

              return {
                autoRemove: false,
                fileInstance: file.fileInstance,
                status: "success",
                shouldUpload: false,
              };
            }}
            draggable={true}
            dragMainText={t("drag_and_drop_files")}
            dragSubText={t("upload_sql_to_generate_diagrams")}
            accept=".sql,.xlsx,.xls"
            onRemove={() => {
              setError({
                type: STATUS.NONE,
                message: "",
              });
              setImportData((prev) => ({ ...prev, src: "", xlsxB64: null }));
            }}
            onFileChange={() =>
              setError({
                type: STATUS.NONE,
                message: "",
              })
            }
            limit={1}
          />
        </TabPane>
      </Tabs>

      <div className="mt-2">
        <Checkbox
          aria-label={t("overwrite_checkbox")}
          checked={importData.overwrite}
          onChange={(e) =>
            setImportData((prev) => ({
              ...prev,
              overwrite: e.target.checked,
            }))
          }
        >
          {t("overwrite_existing_diagram")}
        </Checkbox>
        <div className="mt-2">
          {error.type === STATUS.ERROR ? (
            <Banner
              type="danger"
              fullMode={false}
              description={<div>{error.message}</div>}
            />
          ) : error.type === STATUS.OK ? (
            <Banner
              type="info"
              fullMode={false}
              description={<div>{error.message}</div>}
            />
          ) : (
            error.type === STATUS.WARNING && (
              <Banner
                type="warning"
                fullMode={false}
                description={<div>{error.message}</div>}
              />
            )
          )}
        </div>
      </div>
    </div>
  );
}
