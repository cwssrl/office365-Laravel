<?php

namespace Cwssrl\Office365;

use Cwssrl\Office365\Auth\StaticAccessTokenProvider;
use League\OAuth2\Client\Provider\GenericProvider;
use Microsoft\Graph\Generated\Users\Item\MailFolders\Item\Messages\MessagesRequestBuilderGetRequestConfiguration;
use Microsoft\Graph\GraphServiceClient;
use Microsoft\Kiota\Abstractions\Authentication\BaseBearerTokenAuthenticationProvider;
use Microsoft\Kiota\Abstractions\Serialization\Parsable;
use Microsoft\Kiota\Serialization\Json\JsonSerializationWriter;

class Office365
{
    /** @var GenericProvider */
    private $client;

    public function __construct()
    {
        $this->client = new GenericProvider([
            'clientId'                => config('office365.appId'),
            'clientSecret'            => config('office365.secret'),
            'redirectUri'             => config('office365.redirect_url'),
            'urlAuthorize'            => config('office365.authority') . config('office365.authority_endpoint'),
            'urlAccessToken'          => config('office365.authority') . config('office365.authority_token'),
            'urlResourceOwnerDetails' => 'https://graph.microsoft.com/v1.0/me',
            'scopes'                  => config('office365.scopes'),
        ]);
    }

    public function login()
    {
        return $this->client->getAuthorizationUrl();
    }

    public function getAccessToken(string $code)
    {
        $accessToken = $this->client->getAccessToken('authorization_code', [
            'code' => $code,
        ]);

        return [
            'token'        => $accessToken->getToken(),
            'RefreshToken' => $accessToken->getRefreshToken(),
            'expires'      => $accessToken->getExpires(),
        ];
    }

    public function getUserInfo(string $accessToken)
    {
        $user = $this->graphClient($accessToken)->me()->get()->wait();

        return $this->toArray($user);
    }

    public function getEmails(string $accessToken, int $limit = 10)
    {
        $configuration = new MessagesRequestBuilderGetRequestConfiguration();
        $configuration->queryParameters->orderby = ['receivedDateTime DESC'];
        $configuration->queryParameters->top = $limit;

        $messages = $this->graphClient($accessToken)->me()
            ->mailFolders()
            ->byMailFolderId('inbox')
            ->messages()
            ->get($configuration)
            ->wait();

        return $this->toArray($messages);
    }

    private function graphClient(string $accessToken): GraphServiceClient
    {
        $tokenProvider = new BaseBearerTokenAuthenticationProvider(
            new StaticAccessTokenProvider($accessToken)
        );

        return GraphServiceClient::createWithAuthenticationProvider($tokenProvider);
    }

    private function toArray(?Parsable $model): ?array
    {
        if ($model === null) {
            return null;
        }

        $writer = new JsonSerializationWriter();
        $writer->writeObjectValue(null, $model);

        return json_decode((string) $writer->getSerializedContent(), true);
    }
}
