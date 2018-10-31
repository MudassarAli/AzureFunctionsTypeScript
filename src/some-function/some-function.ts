import { Context, HttpMethod, HttpRequest, HttpResponse, HttpStatusCode } from 'azure-functions-ts-essentials';
import * as auth from "node-sp-auth";
import * as request from "request-promise";

const OBJECT_NAME = 'someObject';
export const TEST_ID = '57ade20771e59f422cc652d9';
export const TEST_REQUEST_BODY: { name: string } = {
  name: 'Azure'
};


const getOne = (id: any) => {
  return {
    status: HttpStatusCode.OK,
    body: {
      id,
      object: OBJECT_NAME,
      ...TEST_REQUEST_BODY
    }
  };
};

const getMany = (req: HttpRequest) => {
  return {
    status: HttpStatusCode.OK,
    body: {
      object: 'list',
      data: [
        {
          id: TEST_ID,
          object: OBJECT_NAME,
          ...TEST_REQUEST_BODY
        }
      ],
      hasMore: false,
      totalCount: 1
    }
  };
};

const insertOne = (req: HttpRequest) => {
  return {
    status: HttpStatusCode.Created,
    body: {
      id: TEST_ID,
      object: OBJECT_NAME,
      ...req.body
    }
  };
};

const updateOne = (req: HttpRequest, id: any) => {
  return {
    status: HttpStatusCode.OK,
    body: {
      id,
      object: OBJECT_NAME,
      ...req.body
    }
  };
};

const deleteOne = (id: any) => {
  return {
    status: HttpStatusCode.OK,
    body: {
      deleted: true,
      id
    }
  };
};


/**
 * Routes the request to the default controller using the relevant method.
 */
/*export function run(context: Context, req: HttpRequest): any {
  let res: HttpResponse;
  const id = req.params
    ? req.params.id
    : undefined;

  switch (req.method) {
    case HttpMethod.Get:
      res = id
        ? getOne(id)
        : getMany(req);
      break;
    case HttpMethod.Post:
      res = insertOne(req);
      break;
    case HttpMethod.Patch:
      res = updateOne(req, id);
      break;
    case HttpMethod.Delete:
      res = deleteOne(id);
      break;

    default:
      res = {
        status: HttpStatusCode.MethodNotAllowed,
        body: {
          error: {
            type: 'not_supported',
            message: `Method ${req.method} not supported.`
          }
        }
      };
  }

  context.done(undefined, res);
}*/

export async function AuthenticateToSharePoint(context: Context, req: HttpRequest) {

  let getAuth = await
    auth
      .getAuth(process.env["SiteUrl"], {
        username: process.env["UserName"],
        password: process.env["Password"],
        domain: process.env["Domain"]
      });

  let headers = getAuth.headers;
  headers['Accept'] = 'application/json;odata=verbose';
  headers['x-forms_based_auth_accepted'] = 'f';
  let requestOpts: any = getAuth.options;
  requestOpts.json = true;
  requestOpts.headers = headers;
  requestOpts.url = process.env["SiteUrl"] + '_api/web';

  let response: any = await request.get(requestOpts)

  return {
    status: 200,
    body: {
      object: response.d.Title
    }
  }
}

/**
 * Routes the request to the default controller using the relevant method.
 */
export function run(context: Context, req: HttpRequest): Promise<any> {
  let res: HttpResponse;
  const id = req.params
    ? req.params.id
    : undefined;

  return AuthenticateToSharePoint(context, req);
}
