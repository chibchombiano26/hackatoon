/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

const awsmobile = {
  aws_project_region: "us-west-2",
  aws_cognito_identity_pool_id: "us-west-2:8c9811f3-48b8-466d-817b-c2d7b9c04506",
  aws_cognito_region: "us-west-2",
  aws_user_pools_id: "us-west-2_SUGA3muLW",
  aws_user_pools_web_client_id: "9rsqpu5ie42e0gtrprfas5ikr",
  oauth: {},
  aws_cognito_username_attributes: ["EMAIL"],
  aws_cognito_social_providers: [],
  aws_cognito_signup_attributes: ["EMAIL"],
  aws_cognito_mfa_configuration: "OFF",
  aws_cognito_mfa_types: ["SMS"],
  aws_cognito_password_protection_settings: {
    passwordPolicyMinLength: 8,
    passwordPolicyCharacters: [],
  },
  aws_cognito_verification_mechanisms: ["EMAIL"],
  aws_appsync_graphqlEndpoint: "https://xoikvv3jr5hrbjfzh5uf2usn3a.appsync-api.us-west-2.amazonaws.com/graphql",
  aws_appsync_region: "us-west-2",
  aws_appsync_authenticationType: "AMAZON_COGNITO_USER_POOLS",
  aws_user_files_s3_bucket: "alucio-beacon-content195645-oavila",
  aws_user_files_s3_bucket_region: "us-west-2",
  aws_dynamodb_all_tables_region: "us-west-2",
  aws_dynamodb_table_schemas: [
    {
      tableName: "ObjectAudit-oavila",
      region: "us-west-2",
    },
  ],

  Auth: {
    identityPoolId: "us-west-2:8c9811f3-48b8-466d-817b-c2d7b9c04506", //REQUIRED - Amazon Cognito Identity Pool ID
    region: "us-west-2", // REQUIRED - Amazon Cognito Region
    userPoolId: "us-west-2_SUGA3muLW", //OPTIONAL - Amazon Cognito User Pool ID
    userPoolWebClientId: "9rsqpu5ie42e0gtrprfas5ikr", //OPTIONAL - Amazon Cognito Web Client ID
  },
  Storage: {
    AWSS3: {
      bucket: "alucio-beacon-content195645-oavila", //REQUIRED -  Amazon S3 bucket name
      region: "us-west-2", //OPTIONAL -  Amazon service region
    },
  },
};

type CreateDocumentFromS3UploadInput = {
  srcFilename: string;
  fileS3Key: string;
};

const createDocumentFromS3Upload = /* GraphQL */ `
  mutation CreateDocumentFromS3Upload($inputDoc: CreateDocumentFromS3UploadInput!) {
    createDocumentFromS3Upload(inputDoc: $inputDoc) {
      id
      tenantId
      documentId
      versionNumber
      srcFilename
      conversionStatus
      status
      srcFile {
        bucket
        region
        key
        url
      }
      srcHash
      srcSize
      pages {
        pageId
        number
        srcId
        srcHash
        isRequired
        speakerNotes
        linkedSlides
      }
      pageGroups {
        id
        pageIds
        name
        locked
      }
      type
      releaseNotes
      changeType
      labelValues {
        key
        value
      }
      customValues {
        fieldId
        values
      }
      title
      shortDescription
      longDescription
      owner
      expiresAt
      hasCopyright
      purpose
      canHideSlides
      distributable
      downloadable
      isInternalGenerated
      semVer {
        minor
        major
      }
      notificationScope
      selectedThumbnail
      publishedAt
      uploadedAt
      uploadedBy
      convertedArchiveKey
      convertedArchiveSize
      convertedFolderKey
      associatedFiles {
        id
        isDistributable
        isDefault
        type
        attachmentId
        status
        createdAt
        createdBy
      }
      editPermissions
      converterVersion
      createdAt
      createdBy
      updatedAt
      updatedBy
      integration {
        externalVersionId
        version
        timestamp
        srcFileHash
        srcDocumentHash
      }
      integrationType
      _version
      _deleted
      _lastChangedAt
    }
  }
`;

import { Amplify, Auth, Storage, API } from "aws-amplify";
const appsync = require("aws-appsync");
import gpl from "graphql-tag";
require("cross-fetch/polyfill");
Amplify.configure(awsmobile);

const docs = [
  {
    thumbnail: "../../assets/doc-1.png",
    title: "NCCN melanoma",
    description: "This is a document",
    url: "https://gist.githubusercontent.com/chibchombiano26/721ae21344f7ca71c3ed184c87aa1146/raw/027085a8d76a6818e2e4a68de2653fed2250a2eb/olympta",
  },
  {
    thumbnail: "../../assets/doc-2.png",
    title: "George Washington Carver",
    description: "This is a document",
    url: "https://gist.githubusercontent.com/chibchombiano26/8aa719d5e61bde40d249a3e49c691591/raw/4e4c54c1c5ec3fa405dea085f9ae680de06975a0/document%25202",
  },
  {
    thumbnail: "../../assets/doc-3.png",
    title: "Document 1",
    description: "This is a document",
    url: "https://gist.githubusercontent.com/chibchombiano26/8aa719d5e61bde40d249a3e49c691591/raw/4e4c54c1c5ec3fa405dea085f9ae680de06975a0/document%25202",
  },
];

let graphClient = undefined;

Office.onReady(async (info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("login-to-beacon").addEventListener(
      "click",
      async (e) => {
        e.preventDefault();
        var user = (<HTMLInputElement>document.getElementById("form2Example11")).value;
        var pass = (<HTMLInputElement>document.getElementById("form2Example22")).value;
        signIn(user, pass);
      },
      true
    );
    document.getElementById("send-to-beacon").addEventListener("click", async () => {
      Office.context.document.getFileAsync(Office.FileType.Compressed, (result) => {
        // jq("#send-to-beacon").attr("disabled", "disabled");
        const button = document.getElementById("send-to-beacon");
        const probressbar = document.getElementById("progress-bar");
        button.setAttribute("disabled", "disabled");
        probressbar.setAttribute("class", "visible");
        // @ts-ignore
        if (result.status == "succeeded") {
          const myFile = result.value;
          myFile.getSliceAsync(0, async (slice) => {
            // @ts-ignore
            if (slice.status == "succeeded") {
              const fileContentArry = slice.value.data;

              const fileContent = new Uint8Array(fileContentArry);
              const file = new File([fileContent], "myFile.pptx", { type: "data:attachment/powerpoint" });
              // calculate progress bar to 100%

              const result = await Storage.put("myFile.pptx", fileContent, {
                level: "private",
                contentType: "data:attachment/powerpoint",
                progressCallback(progress) {
                  console.log(progress);
                },
              });
              console.log(result);
              await createDocumentVersionFromS3(graphClient, result.key, "myFile.pptx");
              button.removeAttribute("disabled");
              probressbar.setAttribute("class", "invisible");
            }
          });

          myFile.closeAsync();
        } else {
          button.removeAttribute("disabled");
          probressbar.setAttribute("class", "invisible");
        }
      });
    });

    // await signIn("alucioqa+oscar-jbuffetto@gmail.com", "TestUser321!");
    generateThumbnails();
  }
});

function generateThumbnails() {
  for (const doc of docs) {
    const node = document.createElement("div");
    node.setAttribute("class", "card w-40 w-100 pt-2");
    node.setAttribute("style", "width: 18rem;");
    const img = document.createElement("img");
    img.src = doc.thumbnail;
    img.setAttribute("class", "card-img-top");
    img.onclick = () => {
      uploadFile(doc.url);
    };

    const cardBody = document.createElement("div");
    cardBody.setAttribute("class", "card-body");
    const title = document.createElement("h5");
    title.setAttribute("class", "card-title");
    title.innerText = doc.title;
    const description = document.createElement("p");
    description.setAttribute("class", "card-text");
    description.innerText = doc.description;
    cardBody.appendChild(title);
    cardBody.appendChild(description);
    node.appendChild(img);
    node.appendChild(cardBody);

    document.getElementById("thumbnails").appendChild(node);
  }
}

export async function run() {
  /**
   * Insert your PowerPoint code here
   */
  const options: Office.SetSelectedDataOptions = { coercionType: Office.CoercionType.Text };

  await Office.context.document.setSelectedDataAsync(" ", options);
  await Office.context.document.setSelectedDataAsync("Hello World!", options);
}

const uploadFile = async (url: string) => {
  const blob = await fetch(url);
  const text = await blob.text();
  PowerPoint.createPresentation(text);
};

const logIn = async (e) => {
  const blob = await fetch(
    "https://gist.githubusercontent.com/chibchombiano26/8aa719d5e61bde40d249a3e49c691591/raw/4e4c54c1c5ec3fa405dea085f9ae680de06975a0/document%25202"
  );
  const text = await blob.text();
  PowerPoint.createPresentation(text);
};

async function signIn(username, password) {
  try {
    const user = await Auth.signIn(username, password);
    graphClient = getGraphQLClientWithJWT(user.signInUserSession.idToken.jwtToken);
    console.log(user);
    document.getElementById("login-form").setAttribute("style", "display: none");
    document.getElementById("contentaddin").setAttribute("style", "display: block");
  } catch (error) {
    console.log("error signing in", error);
  }
}

async function uploadFromS3(fileName: string, srcFilename: string) {
  const updatedTodo = await API.graphql({
    query: createDocumentFromS3Upload,
    variables: {
      inputDoc: {
        fileS3Key: fileName,
        srcFilename: srcFilename,
      },
    },
  });

  console.log(updatedTodo);
}

function getGraphQLClientWithJWT(token) {
  const graphqlClient = new appsync.AWSAppSyncClient({
    url: awsmobile.aws_appsync_graphqlEndpoint,
    region: awsmobile.aws_appsync_region,
    auth: {
      type: "AMAZON_COGNITO_USER_POOLS",
      jwtToken: token,
    },
    disableOffline: true,
  });

  return graphqlClient;
}

const createDocumentVersionFromS3 = async (gpqlClient: any, fileS3Key: string, srcFilename: string) => {
  try {
    const result = (await gpqlClient.mutate({
      mutation: gpl(createDocumentFromS3Upload),
      variables: {
        inputDoc: {
          fileS3Key: fileS3Key,
          srcFilename: srcFilename,
        },
      },
    })) as { data: any };

    return result;
  } catch (e) {
    console.error(e);
    throw e;
  }
};
