using Autodesk.Authentication;
using Autodesk.Authentication.Model;
using Autodesk.Oss;
using Autodesk.Oss.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace Interaction
{
    internal class OssHandler
    {
        private readonly string clientId;
        private readonly string clientSecret;
        private TwoLeggedToken accessToken;
        private readonly OssClient ossClient;
        private readonly string bucketKey;

        public OssHandler(string clientId, string clientSecret)
        {
            this.clientId = clientId;
            this.clientSecret = clientSecret;
            this.bucketKey = GetBucketKey();
            ossClient = new OssClient();
        }

        private async Task GetToken()
        {
            List<Scopes> scopes = new List<Scopes>{
                Scopes.BucketCreate,
                Scopes.BucketRead,
                Scopes.DataRead,
                Scopes.DataWrite,
                Scopes.DataCreate
            };
            accessToken = await new AuthenticationClient().GetTwoLeggedTokenAsync(clientId, clientSecret, scopes);
        }

        private async Task CheckToken()
        {
            if (accessToken == null || accessToken.ExpiresAt <= DateTimeOffset.UtcNow.ToUnixTimeSeconds())
            {
                await GetToken();
            }
        }

        private string GetBucketKey()
        {
            byte[] bytes = SHA256.HashData(Encoding.UTF8.GetBytes(clientId));
            string name = BitConverter.ToString(bytes);
            return $"HandMadeGratingAddin{name}".Replace("-", String.Empty).ToLower();
        }

        public async Task CreateBucket()
        {
            await CheckToken();
            CreateBucketsPayload payload = new CreateBucketsPayload
            {
                BucketKey = bucketKey,
                PolicyKey = PolicyKey.Persistent
            };
            await ossClient.CreateBucketAsync(Region.US, bucketsPayload: payload, accessToken: accessToken.AccessToken);
        }

        public async Task<CreateObjectSigned> GetPresignedUrl(string objectKey, Access access)
        {
            await CheckToken();
            CreateSignedResource resource = new CreateSignedResource
            {
                MinutesExpiration = 30,
                SingleUse = false,
            };
            return await ossClient.CreateSignedResourceAsync(bucketKey, objectKey, resource, accessToken: accessToken.AccessToken, access: access);
        }

        public async Task UploadObjectToBucket(string objectKey, string path)
        {
            await CheckToken();
            using (FileStream reader = File.OpenRead(path))
            {
                await ossClient.UploadObjectAsync(bucketKey, objectKey, reader, accessToken: accessToken.AccessToken);
            }
        }
    }
}
